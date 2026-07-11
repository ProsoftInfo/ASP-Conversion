<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURDisplay.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 03, 2003
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
<!--#include file="../../include/Accpopulate.asp"-->
<%
Dim oDOM,oNodRoot,oNodTemp,oNodDeatils,oNodTaxRoot,oNodEntry,objRs,newElem,newElem1
dim iSno,sDescription,sAmount,sRate,sQty,sValue,sDiscount,dTotal
dim sSalType,sOrgId,sQuery,sPartyName,sRefernceNo
dim sDiscPer,dBasicTotal,dDisTotal
dim sTaxName,sCatCode,sTaxCode,dTax,sTaxMode,sFormula,dTaxValue
dim iTransNo,sOrgName,sBookName,sParType,sParSubType,sParCode,sBookNo,sOrgPartyCode
dim sVouNo,sCallFrm,sFormVal,sRetVal

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objRs  = server.CreateObject("adodb.recordset")

'Response.Write Request.ServerVariables("HTTP_REFERER")

iTransNo=Request("TransNo")
sCallFrm = Request("CallFrm")
sFormVal = Request("hFormVal")



Response.Write iTransNo
'oDOM.load  server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
sRetVal = GetVouchXML(iTransNo)
oDOM.Load server.MapPath(sRetVal)


set oNodRoot=oDOM.documentElement

sVouNo = oNodRoot.attributes.item(1).nodevalue

for each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
				sOrgName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
				sBookName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="Party" then
				sPartyName=oNodEntry.Text
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sRefernceNo=oNodEntry.Attributes.Item(0).nodeValue &"&nbsp; Dt:"&oNodEntry.Attributes.Item(1).nodeValue
			end if
		next
	end if

	'if oNodTemp.nodeName="Details" then
	'	set oNodDeatils=oNodTemp
	'	sVouNo = sVouNo + " Dt " + oNodTemp.attributes.getnameditem("VouDate").value
	'end if

	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
	end if
next

sQuery = "SELECT OrgnPartyCode FROM APP_M_PartyMaster WHERE PartyCode = "&sParCode&" "

Objrs.Open sQuery,Con
IF Not Objrs.Eof Then
	sOrgPartyCode = Objrs(0)
End IF
Objrs.close

Dim sExp,iCounter,TempNode
sExp = "//Details"
Set TempNode = oNodRoot.selectnodes(sExp)
For iCounter = 0 To TempNode.length - 1
	IF TempNode.Item(iCounter).Attributes.getNamedItem("BasicValue").Value <> "" Then
		'Response.Write iCounter
		set oNodDeatils = TempNode.Item(iCounter)
		sVouNo = sVouNo + " Dt " + TempNode.Item(iCounter).attributes.getnameditem("VouDate").value
	End IF
Next

'Response.Write "<p>sCallFrm="&sCallFrm
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/PrintWindow.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<%IF CStr(sCallFrm) = "A" Then %>
	<form method="POST" name="formname" action="PURCHASEVOUCHERS.ASP">
<%Else%>
	<form method="POST" name="formname" action="PURCHASEVOUCHERENTRY.ASP">
<%ENd IF%>

<input type="hidden" name="hFormVal" value="<%=sFormVal%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Voucher&nbsp;
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack" height="7">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
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
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="105">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable">
								  	<tr>
									  	<td align="center">Invoice Details</td>
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
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" >
								  	<tr>
								  		<td align="center">Voucher</td>
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
                            </tr>
                            <tr>
                            <td align="center">
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="center">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable" width="100%">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <span style="cursor: pointer" Title="Month wise Balance" >
                    <p align="center"><font size="4" face="Webdings">?</font>
                    </span>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Daywise Balance"><font size="3" face="Webdings">?</font>
                    </span>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <span style="cursor: pointer" Title="Voucher History">
                    <font size="4" face="Webdings">?</font>
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
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" class="TableOutlineOnly">
                                <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                <tr>
                            <td class="FieldCellsub">Unit </td>
                            <td width="160" class="FieldCellSub"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>
                            <td class="FieldCellSub">Invoice No. - Date</td>
                            <td class="FieldCellSub" width="160">	<span class="DataOnly"><%=sRefernceNo%>&nbsp;</span></td>
                                </tr>
                                <tr>

<!--                        <td class="FieldCellsub">Book Code </td>
                            <td width="160" class="FieldCellSub"><span class="DataOnly"><%=sBookNo%>&nbsp;</span></td>
-->

                            <td class="FieldCellsub">Voucher No & Dt</td>
                            <td width="160" class="FieldCellSub"><span class="DataOnly"><%=sVouNo%>&nbsp;</span></td>

                            <td class="FieldCellSub">Book Name</td>
                            <td class="FieldCellSub" width="160"><span class="DataOnly"><%=sBookName%>&nbsp;</span></td>
                                </tr>
                                <tr>
                            <td class="FieldCellsub">Party Code </td>

                            <td width="160" class="FieldCellSub"><span class="DataOnly"><%=sOrgPartyCode%></span>
                            </td>
                            <td class="FieldCellSub">Party Name</td>
                            <td class="FieldCellSub" width="160"><span class="DataOnly"><%=sPartyName%>&nbsp;</span>
                            </td>
                                </tr>
                                <tr>
                                                    <td class="MiddlePack" colspan="4"></td>
                                </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
                            <tr>
                <td></td>
                <td valign="top" width="100%">
                <div class="frmBody" id="frm2" style="width: 600; height:242;">
            <table border="0" cellspacing="1" class="ExcelTable" width="574">
        <tr>
    <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Item Description - Account Head</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Quantity</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Invoice<br>
    Rate</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2" width="60">Basic<br>
    Value</td>
    <td class="ExcelHeaderCell" align="center" colspan="2">Discount</td>
    <td class="ExcelHeaderCell" align="center" rowspan="2">Nett<br>
    Basic</td>
        </tr>
        <tr>
    <td class="ExcelHeaderCell" align="center" width="25">%</td>
    <td class="ExcelHeaderCell" align="center" width="60">Value</td>
        </tr>
<%

	Dim sAccheadName,oAccNode,iRow
	iRow = 1
	For Each oNodEntry in oNodDeatils.childNodes
		iSno=oNodEntry.Attributes.Item(0).nodeValue
		sDescription=oNodEntry.Attributes.Item(1).nodeValue
		sAmount=oNodEntry.Attributes.Item(2).nodeValue
		sRate=oNodEntry.Attributes.Item(6).nodeValue
		sQty=oNodEntry.Attributes.Item(3).nodeValue &"&nbsp;"&oNodEntry.Attributes.Item(4).nodeValue
		sValue=oNodEntry.Attributes.Item(7).nodeValue
		sDiscPer=oNodEntry.Attributes.Item(8).nodeValue
		sDiscount=oNodEntry.Attributes.Item(9).nodeValue

		dTotal=CDbl(dTotal)+CDbl(sAmount)
		dBasicTotal=CDbl(dBasicTotal)+CDbl(sValue)

		dDisTotal=CDbl(dDisTotal)+CDbl(sDiscount)
		For Each oAccNode in oNodEntry.childNodes
			IF Cstr(oAccNode.nodeName) = "AccHead" Then
				sAccheadName = oAccNode.Attributes.Item(3).nodeValue
			End IF
		Next

%>
    <tr>
    <td class="ExcelSerial" align="center"><%=iRow%></td>
    <td class="ExcelDisplayCell"><%=sDescription%> - <%=sAccheadName%></td>
    <td class="ExcelDisplayCell" align="Left" width="60"><%=sQty%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sRate,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber(sValue,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="25"><%=FormatNumber(sDiscPer ,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right" width="60"><%=FormatNumber( sDiscount,2,,,0)%></td>
    <td class="ExcelDisplayCell" align="Right"><%=FormatNumber(sAmount,2,,,0)%></td>
        </tr>
<%
		iRow = iRow + 1
	next
'end if
'next %>

        <tr>
    <td class="ExcelSerial" align="center"></td>
    <td class="ExcelSerial" align="center" colspan="3"><p align="right"><b>Total</b>&nbsp;&nbsp;</td>
    <td class="ExcelSerial" align="right"><b><%=FormatNumber(dBasicTotal,2,,,0)%></b></td>
    <td class="ExcelSerial" align="center" width="25">    </td>
    <td class="ExcelSerial" align="right" width="60"><b><%=FormatNumber(dDisTotal,2,,,0)%></b></td>
    <td class="ExcelSerial" align="right"><b><%=FormatNumber(dTotal,2,,,0)%></b></td>
        </tr>
        <input type="Hidden" name="hBasicValue" value="<%=dBasicTotal%>">
        <input type="Hidden" name="hDisValue" value="<%=dDisTotal%>">
        <input type="Hidden" name="hAmount" value="<%=dTotal%>">
<%
dim dInvAmount,iRounded

	dInvAmount = oNodTaxRoot.Attributes.Item(0).nodeValue
	iRounded = oNodTaxRoot.Attributes.Item(3).nodeValue
	dInvAmount = CDbl(dInvAmount)
	iRounded = CDbl(iRounded)


'------------------ Commented on 03/05/2004
'By Suresh due to difference in invoice between last screen & this screen
	'dInvAmount = CDbl(dInvAmount + iRounded)
'------------------ End of comment ----------------

	dInvAmount = FormatNumber(dInvAmount,2,,,0)
	For Each oNodEntry in oNodTaxRoot.childNodes
		if cint(oNodEntry.Attributes.Item(6).nodeValue ) >0 then
			sCatCode=oNodEntry.Attributes.Item(0).nodeValue
			sTaxCode=oNodEntry.Attributes.Item(1).nodeValue
			sTaxMode=oNodEntry.Attributes.Item(2).nodeValue
			sFormula=oNodEntry.Attributes.Item(3).nodeValue
			dTaxValue=oNodEntry.Attributes.Item(4).nodeValue
			dTax=oNodEntry.Attributes.Item(5).nodeValue
			sTaxName=oNodEntry.Text

			%>
			<tr>
				<td class="ExcelSerial" align="center"></td>
				<td class="ExcelSerial" align="right" colspan="5"><%=sTaxName%>&nbsp;</td>
				<%if sTaxMode="F" then %>
				<td class="ExcelDisplayCell" align="right"></td>
				<%elseif CStr(dTaxValue) = "0" Then%>
				<td class="ExcelDisplayCell" align="right"></td>
				<%else%>
				<td class="ExcelDisplayCell" align="right"><%=dTaxValue%>&nbsp;%</td>
				<%end if%>
				<td class="ExcelDisplayCell" align="right"><%=dTax%></td>
				    </tr>
			<%
		end	if
	next
%>


        <tr>
        <td class="ExcelSerial" align="center" colspan="3"></td>
    <td class="ExcelSerial" align="right" colspan="4"><b>Invoice Value&nbsp; </b></td>
    <td class="ExcelSerial" align="right"> <%=FormatNumber(dInvAmount,2,,,0)%> </td>
        </tr>
            </table>
                </div>
                </td>
                <td></td>
                            </tr>
                                                        <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                                                        </tr>
								<tr>
								<td align="center" width="5" class="ClearPixel">
                                &nbsp;
								</td>
								<td valign="top" class="FieldCell" height="20">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="130" valign="top">Amount </td>
                            <td>
                                                               <span class="DataOnly"><%=AmountWords(dInvAmount)%></span>
                            </td>
                                    </table>

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
                            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center">
                                                    <input type="submit" value="Done" name="B2" class="ActionButton" >
                                                    <!--input type="button" value="Print" name="B8" onClick="PrintWindow('PRNPurchaseJournal.asp?iTraNo=<%=iTransNo%>')" class="ActionButton" -->
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
</BODY>
</html>
