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
<meta http-equiv="x-ua-compatible" content="IE=edge">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="UnitBookData"><Book/></script>
<script type="application/xml" data-itms-xml-island="1" ID="OutData"><PartyType/></script>
<script type="application/xml" data-itms-xml-island="1" id="AccHeadData">
<account/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="TempXMLData"><Root></Root></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/SalesDivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/VouTransactions.js"></script>
<script src="../../scripts/MiscPaymentsCompat.js"></script>
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
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
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
		 <input type="date" id="ctlVouFromDate" name="ctlVouFromDate" onblur="MinDate()" class="FormElem itms-date-picker" style="width:89px">
	  </td>
	  <td align="left" class="FieldCellSub" >To</td>
      <td align="left" class="FieldCellSub">
		 <input type="date" id="ctlVouToDate" name="ctlVouToDate" onblur="MinDate()" class="FormElem itms-date-picker" style="width:89px">
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
