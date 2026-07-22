<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
	'Program Name				:	IssGetQtyForMixCodes.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	MOHAMMED ASIF
%>
<!-- #include file="../../include/DatabaseConnection.asp" -->
<%
    Dim sTotQty,sMixCodes,sArrMixCode,nRowCnt,rstemp,sQuery,sItemName,sItemCode,sClassCode
    Dim iRowCnt

    set rstemp = Server.CreateObject("ADODB.Recordset")

    sMixCodes = Request.QueryString("MixCodes")
    sTotQty = Request.QueryString("TotQty")
    sItemCode = Request.QueryString("ItemCode")
    sClassCode = Request.QueryString("ClassCode")
    sArrMixCode = Split(sMixCodes,",")

    sQuery = "Select ItemDescription from VWITEM Where Itemcode = "& sItemCode &" and ClassificationCode = "& sClassCode
    rstemp.Open sQuery,con
    if not rstemp.EOF then
        sItemName = trim(rstemp(0))
    end if
    rstemp.Close


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : Direct Issue - Mix Quantity</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="MixData"><MixData Action="Close"></MixData></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/issGetQtyForMixCodesModern.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type=hidden name="hRefNo" value="<%=sMixCodes%>">
<input type=hidden name="hTotQty" value="<%=sTotQty%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"> Mix Wise Quantity Details
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
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item Description&nbsp;</td>
                                            <td ><span class="DataOnly"><%=sItemName%></span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Total Quantity&nbsp;</td>
                                            <td ><span class="DataOnly"><%=sTotQty%></span></td>
                                        </tr>
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm2" style="width: 250; height:140;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center">S.No.</td>
												<td class="ExcelHeaderCell" align="center">MixCode Name</td>
												<td class="ExcelHeaderCell" align="center">Quantity</td>
											</tr>
											<%
											    sQuery = "Select MixCode,MixName from VW_PRD_MixMaster where MixCode in ("& sMixCodes&")"
											    rstemp.Open sQuery,con
											    if not rstemp.EOF then
											        do while not rstemp.EOF
											            iRowCnt = iRowCnt + 1
											            %>
											                <tr>
											                    <td class="ExcelSerial" align="center"><%=iRowCnt%></td>
												                <td class="ExcelDisplayCell" align="left"><%=trim(rstemp(1))%></td>
												                <td class="ExcelInputCell" align="center">
												                    <input type=text name="txtQtyZ<%=trim(rstemp(0))%>" value="0" class="FormElem" style="text-align:right">
												                    <input type=hidden name="hMixNameZ<%=trim(rstemp(0))%>" value="<%=trim(rstemp(1))%>">
												                </td>
												            </tr>
												        <%
											            rstemp.MoveNext
											        loop
											    end if
											    rstemp.Close
											%>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="reset" value="Reset" name="B2" class="ActionButton">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
</HTML>
