<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PaymentRequests.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	April 05, 2010
	'Modified By                :
	'Modified On                :
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim sUnitID,sFrmDate,sTDate,iCnt,sSql,sFromDate,sToDate
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo
	Dim iTotalPages,iTotalRecords,iPrevPage,iNextPage
	Dim sReqType,sPayTo,sStatus,sSentToVoucher

	Dim Objrs,Objrs1

	set Objrs=server.CreateObject ("ADODB.recordset")
	set Objrs1=server.CreateObject ("ADODB.recordset")

	sUnitID = Session("organizationcode")

	Const iPageSize=16
	iPageNo = trim(Request("hPage"))
	if trim(iPageNo) = "" then iPageNo = 1

	iCurrentPage=CInt(Request.Form("hPageSelection"))

	sFromDate ="01/04/"&Mid(GetFromFinYear,3,4)
	sToDate ="31/03/"&Mid(GetToFinYear,3,4)

	sFrmDate = Request("hFromDate")
	sTDate =  Request("hToDate")

	If sFrmDate = "" Then
		sFrmDate = sFromDate
		sTDate = sToDate
	End IF

	sReqType = Trim(Request("hReqType"))
	sPayTo = trim(Request("hPayTo"))
	sStatus = trim(Request("hStatus"))

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/VoucherEntryCore.js"></script>
<script src="../../scripts/BankVoucher.js"></script>
<script src="../../scripts/ReportReminderCompat.js"></script>
<script src="../../scripts/GetPopUpWindowSize.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="SetDate()">
	<form method="POST" name="formname" action="PaymentRequests.asp" >

	<input type=hidden name="hUnitNo" value="<% =sUnitID%>">
	<input type=hidden name="hUnitName" value="<% =session("orgShortName")%>">
	<input type=hidden name="hFromDate" value="<%=sFrmDate%>">
	<input type=hidden name="hToDate" value="<%=sTDate%>">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hReqType" value="<%=sReqType%>">
	<input type="hidden" name="hPayTo" value="<%=sPayTo%>">
	<input type="hidden" name="hReqTypeS" value="">
	<input type="hidden" name="hStatus" value="<%=sStatus%>">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Payment Requests
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
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
<tr>
<td>
<div>
<table class="CollapseBand" cellspacing="0" cellpadding="0">
<tr>
<td valign="center"><a style="width: 1em; height: 1em;" title="" href="#" onclick="return Div_OnClick(idUnprocessed,'',event)" itms_state="0">
<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
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
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
<tr>
<td class="MiddlePack">
</td>
<td class="MiddlePack" colspan="6">
</td>
</tr>

<tr>
	<td class="FieldCellSub">Date From </td>

    <td class="FieldCellSub" valign="middle">
		<input type="text" id="ctlVouFromDate" name="ctlVouFromDate" class="FormElem itms-date-picker" data-itms-datepicker="1" size="10">
	</td>

	<td class="FieldCellSub">To</td>
    <td class="FieldCellSub" valign="middle">
		<input type="text" id="ctlVouToDate" name="ctlVouToDate" class="FormElem itms-date-picker" data-itms-datepicker="1" size="10">
	</td>
</tr>

<tr>
	<td class="FieldCellSub">Requisition Type</td>
	<td class="FieldCellSub">
		<select Name="selReqType" class="FormElem">
			<option value="S">Select a Request Type</option>
    		<option value="H">Hire Purchase</option>
			<option value="L">Loan</option>
			<option value="B">Blank Cheque</option>
			<option value="A">Regular Payment Cheque</option>
			<option value="O">Regular Payment Chash</option>
			<option value="V">Advance</option>
		</Select>
	</td>
</tr>

<tr>
	<td class="FieldCellSub">Pay To</td>
	<td class="FieldCellSub">
		<input type="text" name="txtPayTo" size="55" class="FormElem" maxlength="50">
		<a href="#" onclick="SelMisParty(); return false;">
		<img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a>
	</td>
</tr>
<tr>
	<td class="FieldCellSub">Status</td>
	<td class="FieldCellSub">
		<input type="checkbox" Name="chkCreated" value="010201" class="FormElem">Created
		<input type="checkbox" Name="chkApproved" value="010203" class="FormElem">Approved
	</td>
</tr>
<tr>
<td class="FieldCell"></td>
<td class="FieldCell" colspan="2" align="center">
	<input type="button" value="Go" name="Cmdgo" class="ActionButton" onclick="Validate()">&nbsp;
	<input type="button" value="Reset" name="Cmdreset" class="ActionButton" onclick="ChkReset()">
</td>

</tr>
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
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top">
<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
<table cellspacing="1" class="ExcelTable" width="100%" >
	<td class="ExcelHeaderCell" width="10px" Rowspan="2">S.No.</td>
	<td class="ExcelHeaderCell" Rowspan="2">
		<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Delete Record" width="15px" height="15px" onclick="CheckSubmit('D')">
		</a>
	</td>
	<td class="ExcelHeaderCell" Colspan="2">Requested
	</td>
	<td class="ExcelHeaderCell" Rowspan="2">Pay To
	</td>
	<td class="ExcelHeaderCell" Rowspan="2">Amount To <bR>Pay
	</td>
	<td class="ExcelHeaderCell" Rowspan="2">Reason
	</td>
</tr>
<tr>
	<td class="ExcelHeaderCell">No - Date</td>
	<td class="ExcelHeaderCell">By</td>
</tr>
<%

Dim sQuery,sPartyName,sReqDetail,sRequestedBy,iSno

'isNull(a.RequestFrom,'P')  RequestFrom

sQuery = " select distinct b.PaymentRequestNo,a.RequestedBy,isNull(a.ApprovedBy,0),convert(char,a.PaymentRequestDt,103),"&_
		 " isNull(b.AccUnitPartyType,0),isNull(b.AccUnitPartySubType,0),isNull(b.AccUnitPartyCode,0),B.ToBePaidTo,B.AmountToPay,B.ReasonForPayment"&_
		 " from Acc_T_PaymentRequestHdr a, Acc_T_PaymentRequestDet b where"&_
		 " b.AccountingUnit='"& sUnitID &"' and b.PaymentRequestNo=a.PaymentRequestNo "&_
		 " and B.AmountToPay > isNull(B.AmountPaid,0)"


sQuery = " select distinct b.PaymentRequestNo,a.RequestedBy,isNull(a.ApprovedBy,0),convert(char,a.PaymentRequestDt,103),"&_
		 " isNull(b.AccUnitPartyType,0),isNull(b.AccUnitPartySubType,0),isNull(b.AccUnitPartyCode,0),B.ToBePaidTo,"&_
		 " SUM(isNull(B.AmountToPay,0)),B.ReasonForPayment,b.StatusOfRequestDet,isNull(a.CreatedTransNo,0),a.RequestedFromAppl,isNull(a.PaymentRequestCode,'')"&_
		 " from Acc_T_PaymentRequestHdr a, Acc_T_PaymentRequestDet b where"&_
		 " b.AccountingUnit='"& sUnitID &"' and b.PaymentRequestNo=a.PaymentRequestNo "

'" and B.AmountToPay > isNull(B.AmountPaid,0) and b.StatusOfRequestDet <> '010203' "

'Response.Write "<p><font color=red>sStatus = "&sStatus

If sStatus <> "" Then
	sQuery = sQuery & " and b.StatusOfRequestDet IN ( '"& Trim(sStatus) &"') "
End IF

If sFrmDate <> "" and sTDate <> "" Then
	'sQuery = sQuery & " and convert(char,a.PaymentRequestDt,103) >= convert(char,'"& sFrmDate &"',103) and convert(char,a.PaymentRequestDt,103) <= convert(char,'"& sTDate&"',103) "
End IF
If Trim(sReqType) <> "" and Trim(sReqType) <> "S" Then
	sQuery = sQuery & " and RequestFor='"&sReqType&"' "
End IF
sQuery = sQuery & " Group by b.PaymentRequestNo,a.RequestedBy,isNull(a.ApprovedBy,0),convert(char,a.PaymentRequestDt,103), isNull(b.AccUnitPartyType,0),isNull(b.AccUnitPartySubType,0),isNull(b.AccUnitPartyCode,0),B.ToBePaidTo,B.ReasonForPayment,b.StatusOfRequestDet,a.CreatedTransNo,a.RequestedFromAppl,a.PaymentRequestCode"
'Response.write "<textarea>"& sQuery &"</textarea>"
'Response.Write "<p>sQuery = "& sQuery
With Objrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
End With

set objrs.ActiveConnection = Nothing

If Not objrs.EOF Then
	iTotalPages = objrs.PageCount
	iTotalRecords = objrs.RecordCount
	objrs.AbsolutePage = iPageNo
Else
	iTotalPages = 0
	iTotalRecords = 0

	'iStartRec = 0
	'iEndRec = 0
End If

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
iSno = 1
	Do while Not objrs.EOF and iSno < iPageSize

		sSentToVoucher = "N"

		sQuery = " select PartyName from VwOrgParty where OUDefinitionID='"& sUnitID &"' "&_
				 " and PartyType='"&objRs(4)&"' and PartySubType="&objRs(5)&" and PartyCode="&objRs(6)
		with Objrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs1.ActiveConnection = nothing
		IF Not objRs1.EOF Then
			sPartyName=objRs1(0)
		Else
			sPartyName = objRs(7)
		End IF
		objRs1.Close


		sQuery="select EmployeeName from Ms_EmployeeMaster where EmployeeNumber="&objRs(1)
		with objRs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs1.ActiveConnection = nothing
		IF Not objRs1.EOF Then
			sRequestedBy=objRs1(0)
		End IF
		objRs1.Close

		'To check whether the selected payment is sent to Voucher Table

		sQuery = "Select Distinct BookCode From Acc_t_CreatedVoucherHeader where CreatedTransNo = "& objrs(11)&" and CreatedVouchStatus = '010104' "
		with objRs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs1.ActiveConnection = nothing
		IF Not objRs1.EOF Then
			sSentToVoucher = "Y"
		End IF
		Objrs1.Close
		%>
		<tr>
			<td class="ExcelSerial" align="center"><%=iSno%></Td>
			<td class="ExcelDisplaycell" align="center">
			<%If sSentToVoucher = "Y" Then%>
				<Input type="checkbox" name="chkbox<%=iSno%>" value="<%=Objrs(0)%>:<%=Objrs(10)%>:<%=Objrs(12)%>" disabled>
			<%Else%>
				<Input type="checkbox" name="chkbox<%=iSno%>" value="<%=Objrs(0)%>:<%=Objrs(10)%>:<%=Objrs(12)%>">
			<%End IF%>
			</tD>
			<td class="ExcelDisplaycell" align="center">
			<%If trim(Objrs(10)) = "010203" Then%>
				<Font color=Red><b>*</b></Font>
			<%Else%>
				&nbsp;
			<%End IF%>
			<%=objrs(13)%> - <%=objrs(3)%>
			</td>
			<td class="ExcelDisplaycell" align="Left"><%=Trim(sRequestedBy)%></td>
			<td class="ExcelDisplaycell" align="Left"><%=sPartyName%></td>
			<td class="ExcelDisplaycell" align="Right"><%=FormatNumber(objrs(8),2,,,0)%></td>
			<td class="ExcelDisplaycell" align="Left"><%=objrs(9)%></td>
		</tr>
		<%
		sReqDetail = ""
		iSno = iSno + 1
		objrs.MoveNext
	Loop
objrs.Close

%>
</table>
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5px">
</td>
</tr>
<input type=hidden name="hCnt" value=<%=iSno-1%>>
<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5px" class="ClearPixel">
</td>
<td valign="top" align="right" class="FieldCell"><Font color=Red><b>*</b></Font> Approved Payments
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hPageSelection" value="0">

<%	If iTotalPage >= 2 Then
if iCurrentPage = 1 then
%>
<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
<%		else%>
<input type="button" value=" |< " class="ActionButtonX" onclick="PaginateAcc('1')" id=button3 name=button3>
<input type="button" value=" << " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage - 1%>')" id=button4 name=button4>
<%		end if	%>
<SELECT class="FormElem" onChange="PaginateAcc(this(this.selectedIndex).value)" id=select1 name=select1>
<%
For lnPage = 1 To iTotalPage
If lnPage = iCurrentPage Then
%>
<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotalPage%></OPTION>
<%		else	%>
<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
<%		end if
next
%>
</SELECT>
<%
if iCurrentPage = iTotalPage then
%>
<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

<%		else	%>
<input type="button" value=" >> " class="ActionButtonX" onclick="PaginateAcc('<%=iCurrentPage + 1%>')" id=button7 name=button7>
<input type="button" value=" >| " class="ActionButtonX" onclick="PaginateAcc('<%=iTotalPage%>')" id=button8 name=button8>
<%		end if
End If
%>
</td>
<td align="center" class="ClearPixel" width="5px">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
</td>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td>
	<tr>

		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td class="ActionCell">
                         <input type="button" value="Create" OnClick="Create()" class="ActionButton"  id=button1 name=button1>
                         <input type="button" value="Edit" OnClick="CheckSubmit('E')" class="ActionButton"  id=button2 name=button2>
                         <input type="button" value="Approve" OnClick="CheckSubmit('A')" class="ActionButton"  id=button3 name=button3 >
                         <input type="button" value="Generate Voucher" OnClick="CheckSubmit('G')" class="ActionButtonX"  id=button4 name=button4>
					</td>
				</tr>
			</table>
		</td>

    </tr>
</td>
</tr>

</table>
</td>
<td align="center" class="ClearPixel">
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
