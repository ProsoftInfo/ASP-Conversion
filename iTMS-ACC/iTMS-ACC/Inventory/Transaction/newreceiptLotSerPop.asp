<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	newreceiptLotSerPop.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 20, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<%
dim iItem,iClass,sOrgID,sType,arrTemp,iRecNo,iMRSNo,iSerial,iLot,iCtr,iIssNo,iQtyRec
Dim sQuery,sItemName,sRcptNum,iTotQty,iInvRecNo,iPackNo
iCtr = 0
dim dcrs
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")	

arrTemp = split(trim(Request.QueryString("sTemp")),":")

iRecNo	= arrTemp(1)
iClass	= arrTemp(2)
iItem	= arrTemp(3)
sOrgID	= arrTemp(4)
iMRSNo  = arrTemp(5)
iIssNo  = arrTemp(6)
iLot    = arrTemp(7)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Internal Receipts</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
    sQuery = "Select InvRecNo from APP_T_InternalReceiptHeader where InternalReceiptNo = "& iRecNo
    dcrs.open sQuery,con
    if not dcrs.eof then
        iInvRecNo = dcrs(0)
    end if
    dcrs.close
    
  	sQuery = "Select ItemDescription,ReceiptNumbering from VWItem Where ItemCode = "& iItem
	dcrs.open sQuery,con
	if not dcrs.eof then
	    sItemName = dcrs(0)
	    sRcptNum = dcrs(1)
	end if
	dcrs.close
	
	if trim(iInvRecNo)<>"" then
	    sQuery = "Select SUM(LotQuantityNett) from INV_T_LocationLot where InventoryReceiptNO = " & iInvRecNo &" and ItemCode = "& iItem
	    'Response.write "<textarea cols=50 rows=3 >"& sQuery&"</textarea>"
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        iTotQty = dcrs(0)
	    end if
	    dcrs.close
	else
	    sQuery = "Select SUM(QuantityReturn) from APP_T_InternalReceiptDetails where InternalReceiptNo = "& iRecNo &" and ItemCode = "& iItem
	    dcrs.open sQuery,con
	    if not dcrs.eof then
	        iTotQty = dcrs(0)
	    end if
	    dcrs.close
	end if 'if trim(iInvRecNo)<>"" then
	

%>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot and Serial Details
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
                                            <td class="FieldCell">Item Description</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=sItemName%></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity&nbsp;</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" ><%=iTotQty%></span>
                                            </td>
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
                                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                                      <tr>
                                        <td valign="bottom">
											<table cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td>
														<table cellpadding="0" cellspacing="0" width="100%">
															<tr>
															<td class='GroupTitleLeft' width="30">&nbsp;</td>
															<td class='GroupTitle' width="80">
																<p align="center">&nbsp;Lot &amp; Serial
                                                            </td>
															<td class='GroupTitleRight'><p align="left">&nbsp;</td>
															</tr>
														</table>
                                                    </td>
												</tr>
												<tr>
													<td class=GroupTable>
														<table cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td class=MiddlePack> </td>
															</tr>
															<tr>
																<td class=FieldCellSub> 
																	<div class="frmBody" id="frm6" style="width: 100%; height:178;">
																		<table border="0" cellspacing="1" id="tblLot3" class="ExcelTable" width="100%">
																		
																			<tr>
																				<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
																				<%if trim(sRcptNum)="LS" then %>
																				<td class="ExcelHeaderCell" align="center">Lot - Pack No.</td>
																				<%elseif trim(sRcptNum)="S" then %>
																				<td class="ExcelHeaderCell" align="center">Pack No</td>
																				<%end if%>
																				<td class="ExcelHeaderCell" align="center">Quantity</td>
																			</tr>
																		<%
																		
																		if trim(iInvRecNo)<>"" then
																	        sQuery = "Select LotNumber,PackingNumber,LotQuantityNett from INV_T_LocationLot where InventoryReceiptNO = " & iInvRecNo &" and ItemCode = "& iItem
																	        dcrs.open sQuery,con
																	        if not dcrs.EOF then
																				 do while not dcrs.EOF
																				    iCtr = iCtr + 1
																				    iLot = dcrs(0)
																				    iPackNo = dcrs(1)
																				    iQtyRec = dcrs(2)
																	                %>
            																			
																		                <tr>
																			                <td class="ExcelSerial" align="center"><%=iCtr%></td>
																			                <td class="ExcelDisplayCell" align="center">
																			                <%
																			                    if trim(sRcptNum)="LS" then 
																			                        Response.write iLot &"-"& iPackNo
																			                    elseif trim(sRcptNum)="S" then 
																			                        Response.write iPackNo
																			                    end if
																			                %>
																			                </td>
																			                <td class="ExcelDisplayCell" align="right" width="10">
																				                <input type="text" name="txtQtyA<%=iCtr%>" size="12" value="<%=iQtyRec%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			                </td>
																		                </tr>
																	                <%			
																				    dcrs.MoveNext
																				loop
																			end if 
																		else
																		    iCtr = 0
																		    sQuery = "Select LotNo,PackingNum,QuantityReturn from APP_T_INTERNALRECEIPTDETAILS where InternalReceiptNo ="& iRecNo &" and ItemCode = "& iItem
																		    dcrs.open sQuery,con
																			if not dcrs.eof then
																		        do while not dcrs.eof 
																		            iCtr = iCtr + 1
																		            iLot = dcrs(0)
																				    iPackNo = dcrs(1)
																				    iQtyRec = dcrs(2)
																	            %>
        																			
																		            <tr>
																			            <td class="ExcelSerial" align="center"><%=iCtr%></td>
																			            <td class="ExcelDisplayCell" align="center">
																			            <%
																			                if trim(sRcptNum)="LS" then 
																			                    Response.write iLot &"-"& iPackNo
																			                elseif trim(sRcptNum)="S" then 
																			                    Response.write iPackNo
																			                end if
																			            %>
																			            </td>
																			            <td class="ExcelDisplayCell" align="right" width="10">
																				            <input type="text" name="txtQtyA<%=iCtr%>" size="12" value="<%=iQtyRec%>" maxlength=10 READONLY class="FormElemRead" style="text-align:right">
																			            </td>
																		            </tr>

																	            <%	
																		            dcrs.movenext
																		        loop
																		    end if
																		    dcrs.close
																		end if'if trim(iInvRecNo)<>"" then
																		%>
																		</table>
																	</div>
																</td>
															</tr>
														</table>
                                                    </td>
												</tr>
											</table>
                                        </td>
                                        <td valign="top">
											<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                        </td>
                                        <td valign="top">
											
											</td>
										</tr>
                                    </table>
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
                                                    <input type="button" value="Close" name="B1" class="ActionButton" onClick="window.close()">
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