<%@ Language="VBScript" %>
<% option explicit %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PayamentDistribution.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran R
	'Created On					:	May 04,2013
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
	Session("ACTN")=trim(Request("ACTN"))
%>
<!--#include file="../../include/Databaseconnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!--#include file="../../include/CommonFunctions.asp"-->
<%
Dim rsObj
Dim sQuery,sSelPartyCode,sSelBookNumber,sFinPeriod
Dim sFinFrom,sFinTo,sFromDate,sToDate,sOrgCode,sPartyName
Dim iCnt

set rsObj = Server.CreateObject("ADODB.Recordset")
sOrgCode = session("organizationcode")
sFinPeriod = Session("FinPeriod")
sFinFrom = "01/04/"& split(sFinPeriod,":")(0)
sFinTo = "31/03/"& split(sFinPeriod,":")(1)

sSelPartyCode = Request("hPartyCode")
sSelBookNumber = Request("hBookNumber")
sFromDate = Request("hFromDate")
sToDate = Request("hToDate")

if trim(sFromDate)="" or IsNull(sFromDate) then
    sFromDate = sFinFrom
    sToDate = sFinTo
end if

if trim(sSelPartyCode)<>"" then
    sPartyName = GetPartyName(sSelPartyCode)
end if

	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT SRC="../../scripts/printwindow.js"></SCRIPT>
<SCRIPT SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<Script>
function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function partyRoot(value) {
	if (!value) {
		return null;
	}
	if (value.nodeType === 1) {
		return value;
	}
	if (value.documentElement) {
		return value.documentElement;
	}
	if (value.XMLDocument && value.XMLDocument.documentElement) {
		return value.XMLDocument.documentElement;
	}
	if (typeof value === "string") {
		return new DOMParser().parseFromString(value, "text/xml").documentElement;
	}
	return null;
}

function firstEntry(root) {
	for (var i = 0; root && i < root.childNodes.length; i += 1) {
		if (root.childNodes[i].nodeType === 1 && String(root.childNodes[i].nodeName).toLowerCase() === "entry") {
			return root.childNodes[i];
		}
	}
	return null;
}

function openModernDialog(url, args, features, callback) {
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	} else {
		window.open(url, "_blank", "height=600,width=750,resizable=no,status=no");
	}
}

function dateControl(name) {
	return document.formname.elements[name] || document.getElementById(name);
}

function ShowVouch(iCrTransNo) {
	openModernDialog("BankVouchView_San.asp?TransNo=" + encodeURIComponent(iCrTransNo), "", "dialogHeight:600px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No");
}

function Init() {
	var form = document.formname;
	dateControl("ctlFromDate").setDate(form.hFromDate.value);
	dateControl("ctlToDate").setDate(form.hToDate.value);
	for (var iCnt = 0; iCnt < form.selBank.options.length; iCnt += 1) {
		if (form.selBank.options[iCnt].value === form.hBookNumber.value) {
			form.selBank.selectedIndex = iCnt;
			break;
		}
	}
}

function SelectParty() {
	var sizeInfo = GetWindowSizeForPopup("2").split(":");
	var programName = sizeInfo[0];
	var popupHeight = sizeInfo[1];
	var popupWidth = sizeInfo[2];
	var partyData = " ? ? ";
	var args = window.PartyData || document.getElementById("PartyData");
	var url = "../../Common/" + programName + "?orgID=" + encodeURIComponent(document.formname.hOrgCode.value) + "&Party=" + encodeURIComponent(partyData);
	openModernDialog(url, args, "dialogHeight:" + popupHeight + "px;dialogWidth:" + popupWidth + "px;Status:No", function (outValue) {
		var root = partyRoot(outValue);
		var entry;
		var partyName;
		if (!root || root.getAttribute("Action") === "CLOSE") {
			return;
		}
		entry = firstEntry(root);
		if (entry) {
			partyName = entry.getAttribute("RetField0") || "";
			if (trim(partyName) !== "") {
				document.getElementById("spanPartyName").textContent = partyName;
				document.formname.hPartyCode.value = entry.getAttribute("RetField1") || "";
			}
		}
	});
}

function Search() {
	document.formname.hBookNumber.value = document.formname.selBank.options[document.formname.selBank.selectedIndex].value;
	document.formname.hFromDate.value = dateControl("ctlFromDate").getDate();
	document.formname.hToDate.value = dateControl("ctlToDate").getDate();
	document.formname.submit();
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="Init()">
<%
	Const iPageSize=16
	Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,hCnt

	iCurrentPage=CInt(Request.Form("hPageSelection"))
%>
	<form method="POST" name="formname" action="">
	<input type="hidden" name="hFromDate" value="<%=sFromDate%>">
	<input type="hidden" name="hToDate" value="<%=sToDate%>">
	<input type="hidden" name="hPartyCode" value="">
	<input type="hidden" name="hBookNumber" value="">
	<input type="hidden" name="hOrgCode" value="<%=sOrgCode%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle">
				<p align="center">PAYMENT DISTRIBUTION

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
<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%" border="0">
<tr>
<td class="MiddlePack" colspan="6">
</td>
</tr>
<tr>
<td class="FieldCellSub" style="width:30px">Bank 
</td>
<td class="FieldCellSub" colspan="4">
<select size="1" name="selBank" class="FormElem">
	<option value="S">Select Bank</option>
	<%
	    sQuery = "Select BookNumber,BookName from Acc_R_ApplicableAccountHeads where BookCode = 02"
	    rsObj.open sQuery,con
	    if not rsObj.eof then
	        do while not rsObj.eof 
	            if trim(sSelBookNumber)= trim(rsObj(0)) then
	                Response.write "<option value="& trim(rsObj(0)) &" Selected>"&trim(rsObj(1))&"</option>"
	            else
	                Response.write "<option value="& trim(rsObj(0)) &">"&trim(rsObj(1))&"</option>"
	            end if
	            
	            rsObj.movenext
	        loop
	    end if
	    rsObj.close
	%>
</select>
</td>
</tr>

<tr>
	<td class="FieldCellsub">Date From
	<td class="FieldCellSub" valign="top" colspan="4">
		<% Response.Write InsertDatePicker("ctlFromDate") %>
	&nbsp;To&nbsp;
		<% Response.Write InsertDatePicker("ctlToDate") %>
	</td>

</tr>

<tr>
	<td class="FieldCellSub">Party</td>
	<td colspan="4" class="FieldCellSub">
	    <span id="spanPartyName" class="DataOnly"><%=sPartyName%>&nbsp;</span>&nbsp;
	    <img src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="SelectParty()" style="cursor: pointer;" alt="Click here to select party" />
	</td>
</tr>

<tr>
<td class="FieldCellsub" colspan="5" align="center" >
	<input type="button" value="Go" name="btnGO" class="ActionButton" onclick="Search()">
	<input type="reset" name="Cmdreset" class="ActionButton">
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
<td align="center" class="ClearPixel" width="5">
<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
</td>
</tr>

<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
    <td class="FieldCell" colspan="3" align="center">
        <b>Cheques Issued List From <span id="spanFrom"><%=sFromDate%>&nbsp;</span> To <span id="spanTo"><%=sToDate%>&nbsp;</span></b>
    </td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top">
<table cellspacing="1" class="ExcelTable" width="100%" >
<tr>
<td class="ExcelHeaderCell" align="center" width="10" >S.No.
</td>
<td class="ExcelHeaderCell" align="center" width="10" >
</td>
<td class="ExcelHeaderCell" align="center" >Bank Name
</td>
<td class="ExcelHeaderCell" align="center" >Cheque No
</td>
<td class="ExcelHeaderCell" align="center" >Date
</td>
<td class="ExcelHeaderCell" align="center" >Paid to
</td>
<td class="ExcelHeaderCell" align="center" >Amount
</td>
<td class="ExcelHeaderCell" align="center" >Delivery By
</td>
<td class="ExcelHeaderCell" align="center" >Delivery to
</td>
</tr>
<%
    sQuery = "Select A.CreatedTransNo,A.BookCode,A.BookNumber,AccountHead,BookName,E.AccUnitPartyCode,E.AccUnitPartyType,E.AccUnitPartySubType,"
    sQuery = sQuery & " PayToRecdFrom,VoucherAmount,PartyName,D.BankInstrumentNo,Convert(varchar,D.BankInstrumentDate,103) from Acc_T_CreatedVoucherHeader A,Acc_R_ApplicableAccountHeads B,Acc_T_CreatedVoucherInstrumentDet D,Acc_T_CreatedVoucherDetails E"
    sQuery = sQuery & "  left outer join App_M_PartyMaster C on C.PartyCode = E.AccUnitPartyCode "
    sQuery = sQuery & " where A.BookCode = B.BookCode and A.BookNumber = B.BookNumber and A.CreatedTransNo = D.CreatedTransNo and  E.CreatedTransNo = A.CreatedTransNo and A.BookCode  = 02 and TransactionType = 'BAP' "
    sQuery = sQuery &" and Convert(datetime,VoucherDate,103) between Convert(datetime,'"& sFromDate &"',103) and Convert(datetime,'"& sToDate &"',103) "
    
    if trim(sSelPartyCode)<>"" then
        sQuery = sQuery & " and E.AccUnitPartyCode = "& sSelPartyCode
    end if
    
    if trim(sSelBookNumber)<>"" and trim(sSelBookNumber)<>"S" then
        sQuery = sQuery & " and A.BookNumber = "& sSelBookNumber
    end if
    rsObj.open sQuery,con
    if not rsObj.eof then
        iCnt = 0
        do while not rsObj.eof
            iCnt = iCnt + 1
            %>
                <tr>
                    <td class="excelserial" align="center">
                        <%=iCnt%>
                    </td>
                    <td class="excelDisplayCell" align="center">
                    </td>
                    <td class="excelDisplayCell" align="center">
                        <%=trim(rsObj(4))%>
                    </td>
                    <td class="excelDisplayCell" align="center">
                        <%=trim(rsObj(11))%>
                    </td>
                    <td class="excelDisplayCell" align="center">
                        <%=trim(rsObj(12))%>
                    </td>
                    <td class="excelDisplayCell" align="center">
                        <%
                            if trim(rsObj(5))<>"0" then
                                response.write trim(rsObj(10))
                            else
                                response.write trim(rsObj(8))
                            end if
                        %>
                    </td>
                    <td class="excelDisplayCell" align="center">
                        <a href="#" class="ExcelDisplayLink" onclick="ShowVouch('<%=rsObj(0)%>'); return false;"><%=FormatNumber(trim(rsObj(9)),2)%></a>
                    </td>
                     <td class="excelDisplayCell" align="center">
                    </td>
                     <td class="excelDisplayCell" align="center">
                    </td>
                </tr>
            <%
            rsObj.movenext
        loop
    end if
    rsObj.close
%>
</table>
<!--/div-->
</td>
<td align="center" class="ClearPixel" width="5">
</td>
</tr>
<tr>
<td align="center" class="MiddlePack" colspan="3">
</td>
</tr>

<tr>
<td align="center" width="5" class="ClearPixel">
</td>
<td valign="top" align="right" class="FieldCell">
<input type=hidden name="hCurrentPage" value=<%=iCurrentPage %>>
<input type=hidden name="hCnt" value=<%=iCnt%>>
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
<p align="center">
    <!--<input type="button" class="actionbutton" />-->
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
