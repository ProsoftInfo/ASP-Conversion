<%@ Language=VBScript %>
<%	Option Explicit%>
<%
	'Program Name				:  MiscInvEdit.asp
	'Module Name				:  Common
	'Author Name				:  Ragavendran R
	'Created On					:  July 16,2013
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
Dim rsTemp,rsObj
Dim sOrgID,sOrgName,sCreatedBy,sQuery,sAppCode,sParType,sInvNo
Dim sPartyCode,sPartySubType,sPartyType,sCrePymtNo,sVouDate,sVouAmt,sBnkType,sPayToRcdFrm
Dim sAppRefNo,sAppRefType,sVouNar,sPartyName,sMiscPartyCode,sMiscPartyName,sAdjStatus,sPayFor
Dim iCreatedBy,sChequeNo,sChequeDate


set rsTemp = Server.CreateObject("ADODB.Recordset")
set rsObj = Server.CreateObject("ADODB.Recordset")

sOrgID = Session("organizationcode")
sOrgName = Session("OrgShortName")
iCreatedBy = Session("userid")
sCreatedBy = Session("username")
sAppCode = Request.QueryString("APPCODE")
sInvNo = Request.QueryString("InvNo")
if sAppCode = 2 then
    sParType = "CR"
else
    sParType = "DR"
end if


sQuery = "Select PartyCode,PartySubType,PartyType,CreatedMiscPymtNo,Convert(varchar,VoucherDate,103),VoucherAmount,"
sQuery = sQuery &" BankInstrumentType,PayToRecdFrom,AppRefNo,AppRefType,MiscPartyCode,AdjustAgainstInvoice,PaymentFor,IsNull(ChequeNo,'') as ChequeNo,IsNull(ChequeDate,'') as ChequeDate from Acc_T_MiscPymtRequestHeader where MiscTransNo = "& sInvNo
rsObj.open sQuery,con
if not rsObj.eof then
    sPartyCode = rsObj(0)
    sPartySubType = rsObj(1)
    sPartyType = rsObj(2)
    sCrePymtNo = rsObj(3)
    sVouDate = rsObj(4)
    sVouAmt = rsObj(5)
    sBnkType = rsObj(6)
    sPayToRcdFrm = rsObj(7)
    sAppRefNo = rsObj(8)
    sAppRefType = rsObj(9)
    sMiscPartyCode = rsObj(10)
    sAdjStatus = rsObj(11)
    sPayFor = rsObj(12)
    sChequeNo = rsObj("ChequeNo")
    sChequeDate = rsObj("ChequeDate")
end if
rsObj.close

if trim(sChequeDate)="" or IsNull(sChequeDate) then sChequeDate = date()

sQuery = "Select VoucherNarration from Acc_T_MiscPaymentReqDetails where MiscTransNo = "& sInvNo
rsObj.open sQuery,con
if not rsObj.eof then
    sVouNar = rsObj(0)
end if
rsObj.close

'Response.write "<p>"& sAdjStatus

sPartyName = GetPartyName(sPartyCode)

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
<script src="../../scripts/itms-modern-compat.js"></script>
<Script SRC="../../scripts/RefTypePop.js"></Script>
<Script SRC="../../scripts/MiscInvoiceCompat.js"></Script>
</HEAD>
<%


%>

<SCRIPT SRC="../../scripts/PrintWindow.js"></SCRIPT>
<BODY leftMargin=0 topMargin=0 onload="Init()">
<FORM NAME="formname" method="post" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=Hidden name="hSupplierCode" value="<%=sPartyCode%>">
<input type=Hidden name="hSupplierName" value="<%=sPartyName%>">
<input type=hidden name="hParType" value="<%=sParType%>">

<input type=hidden name="hRefTypeCode" value="">
<input type=hidden name="hRefno" value="">
<input type=hidden name="hRefDate" value="">
<input type=hidden name="hAppCode" value="<%=sAppCode%>">
<input type=hidden name="hMisPartyCode" value="<%=sMiscPartyCode%>">
<input type="hidden" name="hAppRefType" value="<%=sAppRefType%>">
<input type="hidden" name="hAppRefNo" value="<%=sAppRefNo%>">
<input type="hidden" name="hRefNoDate" value="<%=GetRefNoDate(sAppRefType,sAppRefNo)%>">
<input type="hidden" name="hMiscInvNo" value="<%=sInvNo%>">
<input type="hidden" name="hPayFor" value="<%=sPayFor%>">
<input type="hidden" name="hSubType" value="<%=sPartyType%>|<%=sPartySubType%>">
<input type="hidden" name="hChequeDate" value="<%=sChequeDate%>" />

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">

<tr>
		<td align="center" class=PageTitle height="20">
          <p align="center">Miscellaneous Invoice (Amendment)</p>
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
                                                        RefTypePop 26,2
                                                    else
                                                        RefTypePop 27,3
                                                    end if
                                                %>
                                                </select>
                                                &nbsp;<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item (s)" width="11" height="11" onClick="RefType_Click(); return false;"></a>
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
                                            	<span class="dataonly" id="IDPayTo"><%=sPayToRcdFrm%></span> &nbsp;
	                                            &nbsp;<img id="Img1" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" onclick="SelMisParty()" align="middle" alt="Miscellaneous Party" width="10" height="11">
	                                            <input type=text name="txtPayTo" value="" class=FormElem style="text-align:right" <%if trim(sPayToRcdFrm)<>"" then response.write "readonly" %> >
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
                                                &nbsp;&nbsp;<textarea name="txtPayFor" class="FormElem" rows="2" cols="65"><%=sVouNar%></textarea>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Amount</td>
                                            <td class="FieldCell" colspan="3">
                                                <input type=text name=txtAmount value="<%=sVouAmt%>" class="FormElem" size="10" style="text-align:right">
                                                &nbsp;&nbsp;<Input type="checkbox" name="AdjAgainBill" <%if trim(sAdjStatus)="Y" then response.write "Checked" %>>To be Adjusted Against bill[Select If amount to be borne by party]
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
                                                <input type=radio name=radPayThru value="C" <%if trim(sBnkType)="C" then response.write "checked" %> onclick="ChangePaymentMode()" >Cash&nbsp;&nbsp;
                                                <input type=radio name=radPayThru value="D" <%if trim(sBnkType)="D" then response.write "checked" %> onclick="ChangePaymentMode()">Cheque
                                            </td>
                                            <%if trim(sBnkType)="C" then%>
                                             <td class="FieldCell" align="left" id="tdChequeNo" style="display:none">
                                                No: <input type="text" name="txtChequeNo" class="FormElem" value="<%=sChequeNo%>" /> Date: 
                                            </td>
                                            <td class="FieldCell" align="left" id="tdChequeDate" style="display:none">
                                               <%  Response.write InsertDatePicker("ctlChequeDate") %>
                                            </td>
                                            <%else %>
                                                 <td class="FieldCell" align="left" id="td1" style="display:block">
                                                No: <input type="text" name="txtChequeNo" class="FormElem" value="<%=sChequeNo%>" /> Date: 
                                            </td>
                                            <td class="FieldCell" align="left" id="td2" style="display:block">
                                               <%  Response.write InsertDatePicker("ctlChequeDate") %>
                                            </td>
                                            <%end if %>
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
								    <input type=button name=btnClose value="Update" onclick="CheckSubmit()" class="ActionButton">
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
