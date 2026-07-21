<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ViewIssueEntDetail.asp
	'Module Name				:	Inventory (Issue List)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Oct 15,2013
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
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<!--#include file="../../include/CommonFunctions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Issue Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
    Dim rsObj
	Dim sQuery,sOrgCode,sIssueEntryCode,sIssueDate,sIssuedByName,dtCreatedOn,sIssueFrom,sIssuedToType
	Dim sIssuedToCode,sIssuedToSubCode,sIssueType,sMarkPackFlag,sReturnable,sReturnItem,sAppRefType,sAppRefNo
	Dim sIssueTypeCode,sTypeName,sArrRef,sRefName,sRefNoDate,sReceivedBy,sRemarks,sCallFor
	Dim iIssQty,iPickQty,iBalQty
	
	Dim iIssNo,iIssuedBy,iSLNo
	
	set rsObj = Server.CreateObject("ADODB.Recordset")
	
	iIssNo = Request("IssNo")
	sCallFor = Request("CallFor")
	
	
	sQuery = "	Select OrganisationCode,IssueEntryCode,Convert(varchar,IssueDate,103),IssuedBy,CreatedOn,IssueFrom,IssuedToType,IssuedToCode, "&_
	" IssuedToSubCode,IssueType,MarkPackFlag,Returnable,ReturnItem,IsNull(AppRefType,''),AppRefNo,IssueTypeCode,IsNull(MaterialReceivedBy,'N/A'),IsNull(Remarks,'N/A') from INV_T_MaterialIssueHeader "&_
	" where IssueEntryNO = "& iIssNo
	rsObj.open sQuery,con
	if not rsObj.eof then
	    sOrgCode = rsObj(0)
	    sIssueEntryCode= rsObj(1)
	    sIssueDate= rsObj(2)
	    iIssuedBy= rsObj(3)
	    dtCreatedOn= rsObj(4)
	    sIssueFrom= rsObj(5)
	    sIssuedToType= rsObj(6)
	    sIssuedToCode= rsObj(7)
	    sIssuedToSubCode= rsObj(8)
	    sIssueType= rsObj(9)
	    sMarkPackFlag= rsObj(10)
	    sReturnable= rsObj(11)
	    sReturnItem= rsObj(12)
	    sAppRefType= rsObj(13)
	    sAppRefNo= rsObj(14)
	    sIssueTypeCode= rsObj(15)
	    sReceivedBy =  rsObj(16)
	    sRemarks =  rsObj(17)
	end if
	rsObj.close
	'response.write "sAppRefType="& sAppRefType 
	if trim(sAppRefType)<>"" and trim(sAppRefType)<>"0" then
	    sArrRef = split(GetRefNoDate(sAppRefType,sAppRefNo),",")
	    sRefName = sArrRef(0)
	    sRefNoDate = sArrRef(1)
	else
	    sRefName = "None"
	    sRefNoDate = "N/A"
	end if
	

    sIssuedByName = split(GetUserInfo(iIssuedBy),":")(2)
    
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/PrintWindow.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function DeleteDetails(sOrgCode,sIssNo)
    if confirm("Do you want to delete this issue Permanently?") then
        document.formname.action = "mrsIssueDelete.asp?ISSNO="&sIssNo
        document.formname.submit
    end if 
End Function
'***********************************
Function PrintDetails(sOrgCode,sIssNo)
	sTempValues = sOrgCode&":"&sIssNo
	PrintWindow( "../reports/PRNDICreateDetails.asp?sTemp=" + sTempValues)
End Function
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
<%

    Select Case trim(sIssueTypeCode)
        Case "GEN" 
            sTypeName = "General"
        Case "SUB"
            sTypeName = "Subcontract"
        Case "SER"
            sTypeName = "Services"
        Case "JWK"
            sTypeName = "Job Work"
        Case "TRN"
            sTypeName = "Transfer"
        Case "POS"
            sTypeName = "POS Consumption"
    End Select

%>
	<tr>
		<td align="center" height="20" class="PageTitle">
		    <p align="center">Material Issue (<%=sTypeName%>)
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
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td class="FieldCellSub">Issue From</td>
													<td class="FieldCellSub" valign="top">
													    <span class="DataOnly">
													    <%
													    
													        select case sIssueFrom
													            case "AC"
													                Response.Write "Accounts"
													                case "PA"
													                Response.Write "Purchase"
													                case "SA"
													                Response.Write "Sales"
													                case "IN" 
													                Response.Write "Inventory"
													                case "MA"
													                Response.Write "Maintenance"
													                case "PR"
													                Response.Write "PRODUCTION"
													                default
													                Response.Write "Inventory"
													        end select 
													    %>
													    </span>
													</td>

                                                    <td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Issue To</td>
													<td class="FieldCellSub" valign="top">
													    <span class="DataOnly">
														<%
														   Response.write  IssuedToString(sIssuedToType,sIssuedToCode,sIssuedToSubCode)
													    %>
													    </span>
												</tr>
                                                <tr>
                                                   <td class="FieldCellSub">Reference Name</td>
													<td class="FieldCellSub">
														<span class="DataOnly"><%=sRefName%></span>
													</td>
                                                    <td class="FieldCellSub"></td>
                                                   <td class="FieldCellSub">Issue Date</td>
													<td class="FieldCellSub" valign="middle">
													    <span class="DataOnly"><%=sIssueDate%></span>
													</td>
												</tr>

												<tr>
                                                    <td class="FieldCellSub">Reference No - Date</td>
													<td class="FieldCellSub">
														<span class="DataOnly" align=center id="RefNoDate"><%=sRefNoDate%></span>
    												</td>
                                                    <td class="FieldCellSub"></td>
                                                    <td class="FieldCellSub">Created By</td>
														<td class="FieldCellSub">
															<span class="dataonly"><%=sIssuedByName%></span>
														</td>
												</tr>
												<tr>
												    <td class="FieldCellSub">Cost Center</td>
													<td class="FieldCellSub" valign="top">
														<span class="DataOnly">&nbsp;</span>
													</td>
												   	<td class="FieldCellSub" width="2"></td>
													<td class="FieldCellSub" width="75">Acc. Head</td>
													<td class="FieldCellSub">
													    <span class="DataOnly">&nbsp;</span>
													</td>
												</tr>
                                        <tr>
										    <td class="FieldCellSub">Received By</td>&nbsp;
											<td class="FieldCellSub">
												<span class="DataOnly"><%=sReceivedBy%>&nbsp;</span>
											</td>
											<td></td>
											<td class="FieldCellSub" colspan="2">
											    <span class="DataOnly">
											    <%  
											        if trim(sReturnable)="Y" then 
											            response.write "Returnable" 
											            if trim(sReturnItem)="S" then
											                 response.write " (Same)"
											            else
											                response.write " (Difference)"
											            end if
											        else
											            response.write "Non-Returnable" 
											        end if
											    %>
											    </span>
											</td>
										</tr>
										<tr>
										    <td class="FieldCellsub">Remarks</td>

										    <td class="FieldCellSub">
										    <span class="DataOnly"><%=sRemarks%>&nbsp;</span>
										    
										<td class="FieldCellSub"></td>
											<td class="FieldCellSub" colspan="2">
											    <span class="DataOnly">
											        <%  
											            if trim(sIssueType)="M" then 
											                response.write "Marked"
											                if trim(sMarkPackFlag)="N" then
											                    Response.write "(Pick Pack Now)"
											                else
											                    Response.write "(Pick Pack Later)"
											                end if
											            end if
											        %>
											    </span>
											</td>
										</tr>

                                    </table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel"></td>
								<td valign="top" class="FieldCell" width="100%"><center>
                                    <div align="left">
										<table cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;</td>
															<td class='GroupTitle' width="50"><p align="center">Items</td></center>
															<td class='GroupTitleRight'><p align="left">&nbsp;</td>
														</tr>
													</table>
                                                </td>
											</tr>
											<tr>
												<td class=GroupTable><center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="3"> </td>
														</tr>
														<tr>
															<td class=ClearPixel width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
															<td class=FieldCell>
																<DIV class=frmBody id=frm2 style="height:230;">
																	<table  id="tblLot" border="0" cellspacing="1" class="ExcelTable" width="100%">
																		<tr>
																			<td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">Item Description
																			</td>
																			<td class="ExcelHeaderCell" align="center" colspan="3">Quantity Details</td>
																			<td class="ExcelHeaderCell" align="center" rowspan="2">Additional<br> Details</td>
																		</tr>
																	    <tr>
																	        <td class="ExcelHeaderCell" align="center">Issued</td>
																			<td class="ExcelHeaderCell" align="center">Picked</td>
																			<td class="ExcelHeaderCell" align="center">Balance</td>
																	    </tr>
																	    <%
																	      iSLNo = 0  
																	        sQuery= " Select ItemDescription,SUM(QuantityIssued),SUM(QuantityPicked),D.ItemCode,D.ClassificationCode,"&_
																	                " D.ItemAttributes from INV_T_MaterialIssueDetails D join VWItem V on D.ItemCode = V.ItemCode  where IssueEntryNo = "& iIssNo &""&_
																	                " Group By ItemDescription,D.ItemCode,D.ClassificationCode,D.ItemAttributes"
																	                rsObj.open sQuery,con
																	                if not rsObj.eof then
																	                    do while not rsObj.eof
																	                        iSLNo = iSLNo + 1
																	                        iIssQty = rsObj(1)
																	                        iPickQty =rsObj(2)
																	                        iBalQty = cdbl(iIssQty)-cdbl(iPickQty)
																	                        
																	                        %>
																	                        <tr>
																			                    <td class="ExcelSerial" align="center"><%=iSLNo%></td>
																			                    <td class="ExcelDisplayCell" align="center"><%=rsObj(0)%>
																			                    </td>
																			                    <td class="ExcelDisplayCell" align="center">
																			                    <%
																			                        if cdbl(iIssQty)>0 then
																			                            response.write FormatNumber(iIssQty,3,0,0,0)
																			                        else
																			                            response.write "0.000"
																			                        end if
																			                    %>
																			                    </td>
																			                    <td class="ExcelDisplayCell" align="center">
																			                    <%
																			                        if cdbl(iPickQty)>0 then
																			                            response.write FormatNumber(iPickQty,3,0,0,0)
																			                        else
																			                            response.write "0.000"
																			                        end if
																			                    %>
																			                    </td>
																			                    <td class="ExcelDisplayCell" align="center">
																			                    <%
																			                        if cdbl(iBalQty)>0 then
																			                            response.write FormatNumber(iBalQty,3,0,0,0)
																			                        else
																			                            response.write "0.000"
																			                        end if
																			                    %>
																			                    </td>
																			                    <td class="ExcelDisplayCell" align="center" rowspan="2">&nbsp;
																			                    </td>
																		                    </tr>
																		                    <%
																		                    rsObj.movenext
																	                    loop
																	                end if
																	                rsObj.close
																	    %>
																	</table>
																</div>
															</td>
															<td class=ClearPixel width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
														</tr>
														<tr>
															<td class=MiddlePack width="267" colspan="3"></td>
														</tr>
													</table>
                                                </td>
											</tr>
										</table>
                                    </div>
								</td>
								<td align="center" class="ClearPixel"></td>
							</tr>
                            <tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
												<%if Trim(sCallFor)="D" then %>
                                                    <input type="button" value="Delete" name="btnDelete" class="ActionButton" onClick="DeleteDetails('<%=sOrgCode%>','<%=iIssNo%>')">
                                                <%else%>
                                                    <input type="button" value="Print" name="btnPrint" class="ActionButton" onClick="PrintDetails('<%=sOrgCode%>','<%=iIssNo%>')">
                                                <%end if  %>
                                                    <input type="button" value="Close" name="B3" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel" colspan="3">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
