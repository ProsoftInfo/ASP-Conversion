<%@ Language="VBScript" %>
<% option explicit %>
<%
	'Program Name				:	ADVANCEPAYMNTREQUEST.asp
	'Module Name				:	PURCHASE
	'Author Name				:   RAGAVENDRAN R
	'Created On					:   April 07,2011
	'Modified By				:
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
<!--#include file="../../include/Databaseconnection.asp"-->
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include File="../../include/Purpopulate.asp" -->
<!--#include File="../../include/sessionVerify.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
    Dim rsObj,rsTemp
    Dim iRecCtr,iPrevPage,iTotalPages,nPageCtr,iNextPage,iCreateTransNo
    Dim iMiscNo,dVouDate,nVouAmt
    Dim sOrgID,sQuery,iPartyCode,sPartyName,sParType,sAppCode,sAppRefType,sAppRefNo,sAppRefDate
    Dim sFromDate,sToDate,sArrPeriod,sFinPeriod,sRefName,sRefNoDate,sMiscNoDate,sCreatedTransNo
    sOrgID = Session("organizationcode")
    sFinPeriod = Session("FinPeriod")

    set rsObj = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.Recordset")

        iPartyCode = Request.QueryString("ParCode")
        sFromDate = Request.QueryString("FromDate")
        sToDate = Request.QueryString("ToDate")
        sAppCode = Request.QueryString("APPCODE")

        if trim(sAppCode) = "" then
			sAppCode = Request.Form("hAppCode")
        end if

        if sAppCode=2 then
            sParType = "CR"
        else
            sParType = "DR"
        end if


        if Trim(sFromDate)="" then
            sArrPeriod = Split(sFinPeriod,":")
            sFromDate = "01/04/"& sArrPeriod(0)
            sToDate = "31/03/"& sArrPeriod(1)
        end if

    %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<XML ID="OutData">
	<Root/>
</XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/printwindow.js"></SCRIPT>
<script language="VBScript">
'**************************************************
Function Create()
    document.formname.action = "NewMiscInvoice.asp?APPCODE="& document.formname.hAppCode.value
    document.formname.submit
End Function
'**************************
Function EditInv()
    nInvRow = document.formname.hCnt.value
    iSelectCnt = 0
    for iCnt = 1 to nInvRow
        set objChk = eval("document.formname.ChkZ"&trim(iCnt))
        if objChk.checked = true then
            sInvNo = split(objChk.value,":")(0)
            sCrVouNo = split(objChk.value,":")(1)
            iSelectCnt = iSelectCnt + 1
        end if
    next
    if iSelectCnt > 1 or iSelectCnt <1 then
        alert("Select any one Invoice to Edit")
        exit function
    end if
    if trim(sCrVouNo)<>"" and trim(sCrVouNo)<>"0" then
        alert("Voucher is Created Against the Misc. Invoice, So you can not edit this.")
        exit function
    end if
    if iSelectCnt = 1 then
        document.formname.action = "MiscInvEdit.asp?APPCODE="&document.formname.hAppCode.value&"&InvNo="&sInvNo
        document.formname.submit
    end if
End Function
'*************************************************
Function DeleteInv()
    nInvRow = document.formname.hCnt.value
    iSelectCnt = 0
    for iCnt = 1 to nInvRow
        set objChk = eval("document.formname.ChkZ"&trim(iCnt))
        if objChk.checked = true then
            sInvNo = split(objChk.value,":")(0)
            sCrVouNo = split(objChk.value,":")(1)
            iSelectCnt = iSelectCnt + 1
        end if
    next
    if iSelectCnt > 1 or iSelectCnt <1 then
        alert("Select any one Invoice to delete")
        exit function
    end if
    if trim(sCrVouNo)<>"" and trim(sCrVouNo)<>"0" then
        alert("Voucher is Created Against the Misc. Invoice, So you can not delete this.")
        exit function
    end if
    if iSelectCnt = 1 then
        document.formname.action = "MiscInvoiceDelete.asp?InvNo="&sInvNo
        document.formname.submit
    end if
End Function
'******************************************************
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
End function
'**************************************************
Function Validate()
sParCode = document.formname.hSupplierCode.value
sFrmDate = document.formname.ctlFromDate.getDate()
sToDate = document.formname.ctlToDate.getDate()

    document.formname.action = "MISCINVOICES.ASP?ParCode="& sParCode&"&FromDate="&sFrmDate&"&ToDate="& sToDate
    document.formname.submit

End Function
'**************************************************
Function setdate()
    sFromDate = document.formname.hFromDate.value
    sToDate = document.formname.hToDate.value
    if DateDiff("d",sToDate,date)>0 then
        document.formname.ctlFromDate.setdate = sFromDate
        document.formname.ctlToDate.setDate = sTodate
    else
        document.formname.ctlFromDate.setMinDate = sFromDate
        document.formname.ctlToDate.setMaxDate = date
    end if
End Function
</script>
</head>

<body leftmargin="0" topmargin="0"  onload="Setdate()">
	<form method="POST" name="formname" action="">
	<input type=hidden name="hPage" value="">
	<input type=hidden name="hOrgID" value="<%=sOrgID%>">
	<input type=hidden name="hSupplierName" value="<%=sPartyName%>">
	<input type=hidden name="hSupplierCode" value="<%=iPartyCode%>">
	<input type=hidden name="hFromDate" value="<%=sFromDate%>">
	<input type=hidden name="hToDate" value="<%=sToDate%>">
	<input type=hidden name="hParType" value="<%=sParType%>">
	<input type=hidden name="hAppCode" value="<%=sAppCode%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Misc. Invoices
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
									<td align="center" colspan="3" class="MiddlePack" height="7">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>


<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top" width="100%">
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href onclick="Div_OnClick(idUnprocessed,'')" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
</a>
</td>
<td valign="center" class="SubTitle">&nbsp;&nbsp;
</td>
</tr>

</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
<tr>
<td width="100%">
<div id="idUnprocessed" style="width: 575; display: none">
<table cellpadding="0" cellspacing="0">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="7">
</td>
</tr>

<!--<tr>
	<td class="FieldCellSub">	</td>
	<td class="FieldCellSub">    Item Type	</td>
	<td class="FieldCellSub" colspan="2">
    <select size="1" name="selItemType" class="FormElem">
		<%'popSelItemType(sItemType)	%>
	</select>
	</td>
</tr>-->

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
		Date From
	</td>
	<td class="FieldCellSub">

	<% Response.Write InsertDatePicker("ctlFromDate") %>

	</td>
	<td class="FieldCellSub" colspan="2"></td>
	<td class="FieldCellSub" colspan="2">To	</td>
	<td class="FieldCellSub">

	<%Response.Write InsertDatePicker("ctlToDate") %>

	</td>
</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">
	<% if sAppCode=2 then
	        Response.Write "Supplier"
	   else
	        Response.Write "Party"
	   end if
	%>


	</td>
	<td class="FieldCellSub">
	<span class="dataonly" id="idSupplier"></span> &nbsp;
	&nbsp;<img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="Supplierselect()" align="middle" alt="Supplier Selection" width="10" height="11">
	</td>

</tr>

<tr>
	<td class="FieldCell"></td>
	<td class="FieldCell">

        &nbsp;

	</td>
	<td class="FieldCellSub">

        <p align="right">
            <input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">
        </p>
    </td>
	<td class="FieldCellSub" colspan="2"></td>

	<td class="FieldCellSub" colspan="3">
	<input type="button" value="Reset" name="Cmdreset" class="ActionButtonX" onclick="ResetData()" >
	</td>
</tr>

</table>
</div>
</td>
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
<table border="0" cellspacing="1" class="ExcelTable" width="100%">
<tr>
<td class="ExcelHeaderCell" align="center" rowspan=2 >S.No.</td>
<td class="ExcelHeaderCell" align="center" rowspan=2 >
    <img src="../../assets/images/iTMS%20icons/Deleteicon.gif" alt="Click here to delete the Advance Payment Request" onclick="DeleteInv()">
</td>
<%if sAppCode = 2 then %>
    <td class="ExcelHeaderCell" align="center" rowspan=2 >Supplier Name</td>
<%else %>
    <td class="ExcelHeaderCell" align="center" rowspan=2 >Party Name</td>
<%end if %>
<td class="ExcelHeaderCell" align="center" colspan=2>Reference</td>
<td class="ExcelHeaderCell" align="center" colspan=2>Misc. Invoice</td>
</tr>
<tr>
    <td class="ExcelHeaderCell" align="center">Type</td>
    <td class="ExcelHeaderCell" align="center">No - Date</td>
    <td class="ExcelHeaderCell" align="center">No - Date</td>
    <td class="ExcelHeaderCell" align="center">Amount</td>
</tr>
<%
	'Response.Write "<font color=red>"
        sQuery = "Select MiscTransNo,Convert(varchar,VoucherDate,103),VoucherAmount,PartyCode,"&_
        " AppRefType,AppRefNo,AppRefDate,isNull(CreatedMiscPymtNo,MiscTransNo),IsNull(ReceiptNo,0) as CreatedTransNo from Acc_T_MiscPymtRequestHeader where ApplicationCode = "& sAppCode

        if trim(iPartyCode)<>"" then
           sQuery = sQuery & " and PartyCode = "& iPartyCode
        end if
		'Response.Write "<p> " & sQuery
        rsObj.Open sQuery,con
        if not rsObj.EOF then
            do while not rsObj.EOF
                iRecCtr = iRecCtr + 1
                    iMiscNo = rsObj(0)
                    dVouDate =rsObj(1)
                    nVouAmt = rsObj(2)
                    iPartyCode = rsObj(3)
                    sAppRefType = rsObj(4)
                    sAppRefNo = rsObj(5)
                    sAppRefDate = rsObj(6)
                    sRefNoDate = rsObj(7)
                    sCreatedTransNo = rsObj(8)
                    sMiscNoDate = iMiscNo &" - "& dVouDate

                        sQuery = "Select PartyName from VwOrgParty where PartyCode = "& iPartyCode
                        rsTemp.Open sQuery,con
                        if not rsTemp.EOF then
                            sPartyName = trim(rsTemp(0))
                        end if
                        rsTemp.Close
                        if Trim(sAppRefType)<>"" then
							sQuery = "Select ReferenceName from VW_ReferenceTypes where ReferenceEntryNo ="& sAppRefType
							'Response.Write sQuery
							rsTemp.Open sQuery,con
							if not rsTemp.EOF then
							    sRefName = trim(rsTemp(0))
							end if
							rsTemp.Close
						end if'if Trim(sAppRefType)<>"" then

                    %>
                        <tr>
                            <td class="ExcelSerial" align="center"><%=iRecCtr%></td>
                            <td class="ExcelDisplayCell" align="center">
                                <input type=checkbox name="chkZ<%=iRecCtr%>" value="<%=iMiscNo%>:<%=sCreatedTransNo%>">
                            </td>
                            <td class="ExcelDisplayCell" align="left"><%=sPartyName%></td>
                            <td class="ExcelDisplayCell" align="left"><%=sRefName%></td>
                            <td class="ExcelDisplayCell" align="left"><%=sRefNoDate%></td>
                            <td class="ExcelDisplayCell" align="left"><%=sMiscNoDate%></td>
                            <td class="ExcelDisplayCell" align="right"><%=FormatNumber(nVouAmt,2)%></td>
                        </tr>
                    <%
                rsObj.MoveNext
            loop
        end if
        rsObj.Close
%>
</table>
</td>
<td align="center" class="ClearPixel" width="5">
<input type="hidden" name="hCnt" value="<%=iRecCtr%>" >
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top" align="right">

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
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td valign="middle" class="ActionCell">
    <input type="button" value="New Invoice" name="btnRequest" class="ActionButtonX" onclick="Create()">
    <input type="button" value=" Edit " name="btnEdit" class="ActionButtonX" onclick="EditInv()">
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


<%
Function SelFun()
Dim sTemp
	sTemp = Request.Form("hChoice")
	'Response.Write sTemp
	If sTemp = "A" or sTemp = "C" then
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""EDT"">Edit</OPTION>"  &vbcrlf)
		Response.Write ("<OPTION VALUE=""DEL"">Delete</OPTION>" &vbcrlf)
		'Response.Write ("<OPTION VALUE=""APP"">Finalise</OPTION>" &vbcrlf)
	ElseIf sTemp = "D" then
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""EDT"">Edit</OPTION>"  &vbcrlf)
		Response.Write ("<OPTION VALUE=""DEL"">Delete</OPTION>" &vbcrlf)
		'Response.Write ("<OPTION VALUE=""APP"">Finalise</OPTION>" &vbcrlf)
	ElseIf sTemp = "P" then
		Response.Write ("<OPTION VALUE=""CRE"">Create (Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
	Else

		Response.Write ("<OPTION VALUE=""CRE"">Create (Ref. based)</OPTION>" &vbcrlf)
		Response.Write ("<OPTION VALUE=""CRW"">Create (w/t Ref. based)</OPTION>" &vbcrlf)
	End IF

End Function
%>