<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MaterialReceipts.asp
	'Module Name				:	Inventory (Receipt)
	'Author Name				:
	'Created On					:
	'Modified By				:	Ragavendran R
	'Modified On				:   Feb 28,2013
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptItemEntry.asp
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Receipt</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PurchaseCCIDivClick.Js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../scripts/materialReceipts.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 >
<%
	dim dcrs,dcrs1,dcrs2,iCtr,iIntRecNo,dCreDate,sDept,sRefType,sOrgID,sDeptName
	dim sRecFrom,iGRNNo,dGRNDate,iInvoiceNo,dInvoiceDate,sForUnit,sGRNCode,sRecCode,sRecType,sRecName
	dim bFlag,sOrgName,sItemType,sQuery,sOptType,sCallFrom,sInvoiceCode,sUnitID,sRcptType

	bFlag = false
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	sOrgID = Session("organizationcode")
	'sOrgName = trim(Request.Form("hOrgName"))

	sOptType = Request.QueryString("OptType")
	sCallFrom = Request.QueryString("RCPT")
	sRcptType = Request("RCPTTYPE")
	if trim(sOptType)="" then sOptType = "E"

	if trim(sRcptType)="" then sRcptType="GEN"

	if trim(sCallFrom)<>"A" and trim(sCallFrom)<>"" then
	    sRcptType=sCallFrom
	end if



	sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = "& pack(sOrgID)
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sOrgID = dcrs(0)
		sOrgName = trim(dcrs(1))
	end if
	dcrs.Close


''''''''''''''''''''' Paging Declaration ''''''''''''''''''''''''''''''''''''''''
Const iPageSize=22	'How many records to show
Dim iCurrentPage	'Current Page No.
Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
Dim iPageCtr		'Counter
Dim lnPage
iCurrentPage = 0
if Request.Form("hPageSelection") <> "" then iCurrentPage = CInt(Request.Form("hPageSelection"))

con.CursorLocation = 3

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>

<form method="POST" name="formname" action="">

<input type="hidden" name="hUnitID" value="<%=sOrgID%>">
<input type="hidden" name="hOptType" value="<%=sOptType%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Material Receipts
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class="TabBodyWithTopLine">
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
								<td>
    								<table border="0" cellpadding="0" cellspacing="0" width="100%" class=ExcelTable>
	    							    <tr>
		            						<td>
                                                 <div align="left">
                                                    <table class="CollapseBand" cellspacing="0" cellpadding="0" width=100% >
											            <tr>
												            <td valign="middle"><a style="width: 1em; height: 1em;"  onclick="Div_OnClick(idUnprocessed,'');">
													            <img id="ImgSearch" style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
													            </a>
												            </td>
												            <td valign="middle"class="SubTitle">&nbsp;&nbsp;
												                <%if sOptType="E" then %>
												                    <input type=radio name=radReceipts value="E" class="FormElem" onclick="Submit(this)" checked > External
												                    <input type=radio name=radReceipts value="I" class="FormElem" onclick="Submit(this)"> Internal
												                <%else%>
												                    <input type=radio name=radReceipts value="E" class="FormElem" onclick="Submit(this)"> External
												                    <input type=radio name=radReceipts value="I" class="FormElem" onclick="Submit(this)" checked> Internal
												                <%end if%>
												            </td>
											            </tr>
										            </table>
							                        <div align="left" id="idUnprocessed" style="display: none;">
						                                <table cellpadding="0" cellspacing="0" border="0" width="100%" class="BodyTable">
						                                 <%if sOptType="E" then %>
						                                    <tr>
						                                        <td class="FieldCellSub">Receipt Type</td>
						                                        <td class="FieldCellSub">
						                                            <select id="cmbRcptType" class="FormElem">
						                                                <%
						                                                    sQuery = "Select ReceiptIssueTypeCode,ReceiptIssueTypeDesc from APP_M_ReceiptIssueTypes where ApplicableFor in ('B','R')"
						                                                    dcrs.open sQuery,con
						                                                    if not dcrs.eof then
						                                                        do while not dcrs.eof
						                                                            if trim(sRcptType)=trim(dcrs(0)) then
						                                                                response.write "<option value="& trim(dcrs(0)) &" selected>"& trim(dcrs(1)) &"</option>"
						                                                            else
						                                                                response.write "<option value="& trim(dcrs(0)) &">"& trim(dcrs(1)) &"</option>"
						                                                            end if
						                                                            dcrs.movenext
						                                                        loop
						                                                    end if
						                                                    dcrs.close
						                                                %>
						                                            </select>
						                                        </td>
						                                    </tr>
						                                <%end if 'if sOptType="E" then %>
						                                    <tr>
								                                <td class="FieldCellSub"></td>
								                                <td class="FieldCell">
									                                <input type="button" name="ButGo" value="Go" class="ActionButton" onclick="SubmitMe()">
									                                <!--<input type="button" name="btnCreateInv" value="Create Invoice" class="ActionButtonX" onclick="CreateInvoice()">-->
								                                </td>
							                                </tr>
						                                </table>
						        	                </div>
                                                </div>
				            				</td>
					    			    </tr>
						    		</table>
								</td>
								 <td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>

							<%if sOptType = "E" then %>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>

                            <tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm5" style="height:360;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="30%">Received From</td>
												<td class="ExcelHeaderCell" align="center">Receipt Number</td>
												<td class="ExcelHeaderCell" align="center">Receipt Type</td>
												<td class="ExcelHeaderCell" align="center">Invoice No. / Date</td>
												<!--td class="ExcelHeaderCell" align="center">Recd. for Unit</td-->
											</tr>
										<%

											sQuery = "SELECT DISTINCT RECEIPTNUMBER FROM RCV_T_ACTUALRCPTITEMDET WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND RECEIPTNUMBER IN (SELECT DISTINCT RECEIPTNUMBER FROM RCV_T_ACTUALRECEIPTHEADER WHERE ReceiptNumber is not null "
										'	if Trim(sCallFrom)="JWK" then
										'		sQuery = sQuery & " and ReceiptAs=2"
										'	elseif Trim(sCallFrom)="SUB" then
										'		sQuery = sQuery & " and ReceiptAs=3"
										'	elseif Trim(sCallFrom)="TR" then
										'		sQuery = sQuery & " and ReceiptAs=7"
										'	end if

										    if trim(sRcptType)<>"" then
										        sQuery = sQuery & " and ReceiptAs ='"& sRcptType &"'"
										    end if
											sQuery = sQuery & ") ORDER BY RECEIPTNUMBER Desc"
											'Response.write "<textarea>"& sQuery &"</textarea>"

											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
												'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   													dcrs.PageSize = iPageSize
													If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
													dcrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
													iTotPage = dcrs.PageCount					'stores total no. of pages
												'''''''''''''''''''''''''''''''''''''''''''''''''''''''
													For iPageCtr = 1 to dcrs.PageSize

													sInvoiceCode = ""
													iInvoiceNo = ""
													dInvoiceDate = ""
													bFlag = false

													sQuery = "Select PARTYNAME,G.GRNNUMBER,CONVERT(CHAR,G.GRNDATE,103),ORGUNITSHORTDESCRIPTION,GRNCODE,RECEIPTCODE,0,ISNULL(TRANSFEREDFROM,'-'),A.ReceiptAs from RCV_T_GateReceiptHeader G,RCV_T_ActualReceiptHeader A,App_M_PartyMaster P,DCS_OrganizationUnitDefinitions D where A.GRNNumber = G.GRNNUmber and P.PartyCode = G.PartyCode and G.ReceivedByUnit = D.OUDefinitionID and A.ReceiptNumber = "&  dcrs(0)
												'	Response.Write "<textarea>"& sQuery &"</textarea>"
													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = sQuery
														.ActiveConnection = con
														.Open
													end with

													set dcrs1.ActiveConnection = nothing
													if not dcrs1.EOF then
														sRecFrom = trim(dcrs1(0))
														iGRNNo = trim(dcrs1(1))
														dGRNDate = trim(dcrs1(2))
														sForUnit = trim(dcrs1(3))
														sGRNCode = trim(dcrs1(4))
														sRecCode = trim(dcrs1(5))
														sRecType = trim(dcrs1(8))

															if trim(dcrs1(6)) = "7" then
																with dcrs2
																	.CursorLocation = 3
																	.CursorType = 3
																	.Source = "SELECT ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(trim(dcrs1(9))) & ""
																	.ActiveConnection = con
																	.Open
																end with
																set dcrs2.ActiveConnection = nothing
																if not dcrs2.EOF then
																	sRecFrom = trim(dcrs2(0))
																end if
																dcrs2.Close
															end if

															if trim(sRecType)<>"" then

															        sRecName = GetRcptIssName(sRecType)

															end if

													end if
													dcrs1.Close

													sQuery = "Select InvoiceNumber,InvoiceCode,Convert(varchar,SuppInvoiceDate,103) from RCV_T_InvoiceHeader where AppRefType = 8 and AppRefNo = "& dcrs(0)
													dcrs1.Open sQuery,con
													if not dcrs1.EOF then
														iInvoiceNo = Trim(dcrs1(0))
														sInvoiceCode = Trim(dcrs1(1))
														dInvoiceDate = Trim(dcrs1(2))
													end if
													dcrs1.Close

													with dcrs1
														.CursorLocation = 3
														.CursorType = 3
														.Source = "SELECT DISTINCT ITEMCODE FROM RCV_T_ACTUALRCPTITEMDET WHERE RECEIPTNUMBER = " & trim(dcrs(0)) & ""
														.ActiveConnection = con
														.Open
													end with
													set dcrs1.ActiveConnection = nothing
													if not dcrs1.EOF then
														do while not dcrs1.EOF
															with dcrs2
																.CursorLocation = 3
																.CursorType = 3
																.Source = "SELECT ITEMCODE FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & trim(dcrs1(0)) & ""
																.ActiveConnection = con
																.Open
															end with
															set dcrs2.ActiveConnection = nothing
															if not dcrs2.EOF then
																bFlag = false
															else
																bFlag = true
																dcrs2.Close
																exit do
															end if
															dcrs2.Close
														dcrs1.MoveNext
														loop
													else
														bFlag = true
													end if
													dcrs1.Close
													if not bFlag then
														iCtr = iCtr + 1
										%>
													<tr>
														<td class="ExcelSerial" align="center"><%=iCtr%></td>
														<td class="ExcelDisplayCell"><%=sRecFrom%></td>
														<td class="ExcelDisplayCell" align="center">
															<a href="receiptItemEntry.asp?rcptNo=<%=trim(dcrs(0))%>&sOrg=<%=sForUnit%>&gDate=<%=FormatDate(dGRNDate)%>" class="ExcelDisplayLink"><%=sRecCode%></a>
														</td>
														<td class="ExcelDisplayCell"><%=sRecName%></td>
														<td class="ExcelDisplayCell" align="center"><%=sInvoiceCode%> - <%=dInvoiceDate%></td>
														<!--td class="ExcelDisplayCell"><%=sForUnit%></td-->
													</tr>
										<%
													end if
												dcrs.MoveNext
												if dcrs.EOF then exit for
												Next
											end if
											dcrs.Close
										%>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <%elseif trim(sOptType)="I" then %>

                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>

							<tr>
								<td align="center"></td>
								<td>
									<div class="frmBody" id="frm3" style="height:360;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="30%">Received From</td>
												<td class="ExcelHeaderCell" align="center">Int. Rcpt No. - Date</td>
												<td class="ExcelHeaderCell" align="center">Reference Type</td>
												<td class="ExcelHeaderCell" align="center">Reference No. - Date</td>

											</tr>
										<%
										    Dim sRefValues,sArrRefValues,sRefName,sRefCode,sRefDate
										    Dim sAppRefType,sAppRefNo
											''Changed by Ragav on Aug 23,2012 becuase while creating internal receipt directly accounting is doing so here displaying all the internal receipt
											'sQuery = "SELECT DISTINCT INTERNALRECEIPTNO,CONVERT(CHAR,CREATEDON,103),CREATEDFROMDEPT,REFTYPE = CASE REFTYPE WHEN 'R' THEN 'Repacking' WHEN 'M' THEN 'MR / Direct Issue' WHEN 'C' THEN 'CCI Release' WHEN 'F' THEN 'Manufactured' WHEN 'N' THEN 'None' END,ORGANISATIONCODE FROM APP_T_INTERNALRECEIPTHEADER WHERE STATUS = 'N' AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND REFTYPE IN ('M','N','F','C','R') ORDER BY INTERNALRECEIPTNO DESC"
											'sQuery = "SELECT DISTINCT INTERNALRECEIPTNO,CONVERT(CHAR,CREATEDON,103),CREATEDFROMDEPT,REFTYPE = CASE REFTYPE WHEN 'R' THEN 'Repacking' WHEN 'M' THEN 'MR / Direct Issue' WHEN 'C' THEN 'CCI Release' WHEN 'F' THEN 'Manufactured' WHEN 'N' THEN 'None' END,ORGANISATIONCODE FROM APP_T_INTERNALRECEIPTHEADER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND REFTYPE IN ('M','N','F','C','R') ORDER BY INTERNALRECEIPTNO DESC"
											'' Chnaged by Ragav on Feb 28,2013
											sQuery = "SELECT DISTINCT INTERNALRECEIPTNO,CONVERT(CHAR,CREATEDON,103),CREATEDFROMDEPT,REFTYPE,ORGANISATIONCODE,IsNull(AppRefType,0),AppRefNo,AppRefDate FROM APP_T_INTERNALRECEIPTHEADER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " AND REFTYPE IN ('M','N','F','C','R') ORDER BY INTERNALRECEIPTNO DESC"
											'Response.Write "<textarea>"& sQuery &"</textarea>"
											with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with

											'Response.Write "<p>" & dcrs.Source
											set dcrs.ActiveConnection = nothing
											if not dcrs.EOF then
												'''''''''''''''''''''''''''''''''''''''''''''''''''''''
   													dcrs.PageSize = iPageSize
													If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
													dcrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
													iTotPage = dcrs.PageCount					'stores total no. of pages
												'''''''''''''''''''''''''''''''''''''''''''''''''''''''
													For iPageCtr = 1 to dcrs.PageSize
												sDeptName=""
													iCtr = iCtr + 1
													iIntRecNo = dcrs(0)
													dCreDate = dcrs(1)
													sDept = dcrs(2)
													sRefType = dcrs(3)
													sUnitID = dcrs(4)
													sAppRefType = dcrs(5)
													sAppRefNo = dcrs(6)

												'	if Trim(sDeptName)="" then
												'		sQuery="Select IssuedForDescription from Inv_M_IssuedFor where IssuedForCode = '"& sDept &"'"
												'		dcrs1.Open sQuery,con
												'		if not dcrs1.EOF then
												'			sDeptName = Trim(dcrs1(0))
												'		end if
												'		dcrs1.Close
												'	end if
												sRefName = ""
													if Trim(sDeptName)="" then
														sQuery="Select DepartmentName from App_M_Departments where DeptShortName = '"& sDept &"'"
														dcrs1.Open sQuery,con
														if not dcrs1.EOF then
															sDeptName = Trim(dcrs1(0))
														end if
														dcrs1.Close
													end if


													if trim(sAppRefType)<>"0" and trim(sAppRefType)<>"" and trim(sAppRefNo)<>"0" then
													    sRefValues = GetInfoRefType(sAppRefType,sAppRefNo,sUnitID)
													    sArrRefValues = split(sRefValues,":")

													    sRefName = sArrRefValues(0)
													    sRefCode = sArrRefValues(1)
													    sRefDate= sArrRefValues(2)
													else
													    sRefName = "None"
													    sRefCode = ""
													    sRefDate = ""
													end if

										%>
													<tr>
														<td class="ExcelSerial" align="center"><%=iCtr%></td>
														<td class="ExcelDisplayCell"><%=sDeptName%></td>
														<td class="ExcelDisplayCell" align="center">
															<a href="newreceiptItemEntry.asp?rcptNo=<%=trim(iIntRecNo)%>&sOrg=<%=trim(sUnitID)%>&iDate=<%=dCreDate%>" class="ExcelDisplayLink"><%=trim(iIntRecNo)%> - <%=trim(dCreDate)%></a>
														</td>
														<td class="ExcelDisplayCell"><%=sRefName%></td>
														<td class="ExcelDisplayCell"><%=sRefCode%>-<%=sRefDate%></td>
													</tr>
										<%
												dcrs.MoveNext
												if dcrs.EOF then exit for
												next
											end if
											dcrs.Close
										%>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>

							<%end if 'if sOptType = "E" then %>
							<tr>
								    <td align="center" class="ClearPixel">
								        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
								    <td align="left" class="FieldCell">
								        <table width="100%">
								            <tr>
								                <td class="FieldCell">
								                    <%if sCallFrom="IS" then%>
							    	                &nbsp;&nbsp;<font color=Red>*</font> Indicate MR based Issue
							                        <%end if  %>
								                </td>
								                <td class="FieldCell" align="right">
									                    <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
                                                        <Input Type=Hidden name="hCtr" Value="<%=iCtr%>" >
                                                        <Input Type=Hidden name="hPageSelection" Value="" >

										                <%	If iTotPage >= 2 Then
												                if iCurrentPage = 1 then
										                %>
										                <input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
										                <input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
										                <%		else	%>
										                <input type="button" value=" |< " class="ActionButtonX" onclick="Paginate('1')" id=button3 name=button3>
										                <input type="button" value=" << " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage - 1%>')" id=button4 name=button4>
    									                <%		end if	%>
    									                <SELECT class="FormElem" onChange="Paginate(this(this.selectedIndex).value)" id=select1 name=select1>
    									                <%
											                For lnPage = 1 To iTotPage
												                If lnPage = iCurrentPage Then
										                %>
											                <OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotPage%></OPTION>
										                <%		else	%>
											                <OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
    									                <%		end if
    										                next
    									                %>
    									                </SELECT>
    									                <%
    											                if iCurrentPage = iTotPage then
    									                %>
										                <input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
										                <input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

    									                <%		else	%>
										                <input type="button" value=" >> " class="ActionButtonX" onclick="Paginate('<%=iCurrentPage + 1%>')" id=button7 name=button7>
										                <input type="button" value=" >| " class="ActionButtonX" onclick="Paginate('<%=iTotPage%>')" id=button8 name=button8>
    									                <%		end if
											                End If
										                %>
								                </td>
								            </tr>
								        </table>
									</td>
									<td align="center" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								    </td>
								</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                    <!--<input type="button" value="OK" name="B1" class="ActionButton" onClick="window.location.href='receiptEntry.asp'">-->
                                                    <input type="button" value="New Internal Receipt" name="btnnew" class="ActionButtonX" onClick="CreateReceipt('<%=sOrgID%>')">
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
