<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	IssuePickedDetails.asp
	'Module Name				:	Inventory (Picked Details)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 27,2012
	'Modified By                :   
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/Populate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>MR Pick Issue - Item List</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<%
    Dim rsPick,rsObj,rsTemp
    Dim sQuery,sOrgID,sOrgName,sOptionName,sItemDesc,sUserName
    Dim iPickNo,dtPick,iCtr,iTotPickQty,iIssueEntryNo,dtIssueDate,sIssueCode,sPickedBy
    
    set rsPick = Server.CreateObject("ADODB.Recordset")
    set rsObj = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.RecordSet")
    
    sOrgID = session("organizationcode")
    sUserName = getUserID
    
    Response.write "<font color=red>"
    
    iPickNo = Request("PickNo")
	
	sQuery = "Select PickNumber,Convert(varchar,PickedOn,103),IssueEntryNo,PickedBy from INV_T_IssuePick where PickNumber = "&iPickNo
	rsPick.open sQuery,con
	if not rsPick.eof then
	    dtPick = trim(rsPick(1))
	    iIssueEntryNo = rsPick(2)
	    sPickedBy = rsPick(3)
	end if
	rsPick.close
	
	sQuery = "Select OrgUnitDescription from DCS_OrganizationUnitDefinitions where OUDefinitionID = "& sOrgID 
	rsObj.open sQuery,con
	if not rsObj.eof then
		sOrgName = trim(rsObj(0))
	end if
	rsObj.close 
	
    sUserName = trim(session("username"))
	
	sQuery = "Select IsNull(IssueEntryCode,IssueEntryNo),Convert(varchar,IssueDate,103) from INV_T_MaterialIssueHeader where IssueEntryNo = "& iIssueEntryNo
	rsObj.open sQuery,con
	if not rsObj.eof then
	    sIssueCode = rsObj(0)
	    dtIssueDate = rsObj(1)
	end if
	rsObj.close
	
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/issuePickedDetailsModern.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action>
<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Picked Item Details
		</td>
	</tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>	
		<td valign="top">
			<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >			
				<tr>								
					<td class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">						
						    <tr>
						        <td align="center">
						        </td>
								<td width="100%" colspan="4">
								    <div align="left">
									    <table border="0" cellspacing="0" cellpadding="0" >
								            <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
							                        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
						                        </td>
					                        </tr>
									        <tr>
											    <td class="FieldCellSub">Pick Number&nbsp;</td>
											    <td class="FieldCell"><span class="DataOnly"><%=iPickNo%>&nbsp;</span></td>
											    <td class="FieldCellSub">Pick Date</td>
											    <td class="FieldCell"><span class="DataOnly"><%=dtPick%>&nbsp;</span></td>
											</tr>
											<tr>
											    <td class="FieldCellSub">Unit Name&nbsp;</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>
											    <td class="FieldCellSub">Pick By</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sUserName%>&nbsp;</span></td>
											</tr>
											 <tr>
											    <td class="FieldCellSub">Issue Number&nbsp;</td>
											    <td class="FieldCell"><span class="DataOnly"><%=iPickNo%>&nbsp;</span></td>
											    <td class="FieldCellSub">Issue Date</td>
											    <td class="FieldCell"><span class="DataOnly"><%=dtPick%>&nbsp;</span></td>
											</tr>
									    </table>
									</div>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="5" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr></tr>
							<tr>
								<td align="center">
								</td>
								<td width="100%" colspan="4">
								    <div style="width:90%;">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%"> 
										    <tr>
											    <td>
												    <div class="frmBody" id="frm3" style="width: 100%; height:390">
													    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
														    <tr>
															    <td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
															    <td class="ExcelHeaderCell" align="center">Item Description</td>
															    <td class="ExcelHeaderCell" align="center" width="20%">Picked Quantity</td>
														    </tr>
													    <%
													        Response.Write "<font color=red>"
													     '   sQuery = "Select PickNumber,IP.Itemcode,IP.ClassificationCode,QuantityPicked,L.LotNumber,L.PackingNumber "
													      '  sQuery = sQuery & " from INV_T_IssuePickDetails IP Left Outer Join Inv_T_LocationLot L on IP.SerialNo = L.SerialNumber"
													       ' sQuery = sQuery & " and IP.ItemCode=L.ItemCode and IP.ClassificationCode = L.ClassificationCode and PickNumber = "& iPickNo
													       ' sQuery = sQuery & " Order By IP.ItemCode "
													        
													        sQuery = "Select ItemDescription,IP.ItemCode,IP.ClassificationCode,IsNull(ItemAttributes,''),isNull(SUM(QuantityPicked),0) from "
													        sQuery = sQuery & " INV_T_IssuePickDetails IP,VWItem V where IP.ItemCode=V.ItemCode and PickNumber = "& iPickNo
													        sQuery = sQuery & " Group By ItemDescription,IP.ItemCode,IP.ClassificationCode,ItemAttributes  Order By IP.ItemCode"
													        rsObj.open sQuery,con
													        if not rsObj.eof then
													            do while not rsObj.eof
													                iCtr = iCtr + 1
													                sItemDesc = trim(rsObj(0))
													                iTotPickQty = rsObj(4)
													                
													                if trim(rsObj(3))<>"" then
													                    sQuery= "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&rsObj(3)
													                    rsTemp.open sQuery,con
													                    if not rsTemp.eof then
													                        sOptionName = trim(rsTemp(0))
													                    end if
													                    rsTemp.close
													                end if
													                
													                if trim(sOptionName)<>"" then
													                    sItemDesc = sItemDesc & " [ "&sOptionName&" ]"
													                end if
													                
													                %>
													                    <tr>
													                        <td class="ExcelDisplayCell"><%=iCtr%></td>
													                        <td class="ExcelDisplayCell"><%=sItemDesc%></td>
													                        <td class="ExcelDisplayCell" align="right">
													                            <a href="#" class="ExcelDisplayLink" onClick="DisplayPack('<%=iPickNo%>','<%=rsObj(1)%>','<%=rsObj(2)%>','<%=rsObj(3)%>')"><%=FormatNumber(iTotPickQty,2)%></a>
													                        </td>
													                    </tr>
													                <%
													                rsObj.MoveNext
													            loop
													        end if
													        rsObj.close
													    %>

													    </table>
												    </div>
											    </td>
										    </tr>
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
												<input type="button" value="Issue" name="issue" class="ActionButton" onClick="CheckSubmit()">
 												<input type="button" value="Cancel" name="cancel" class="ActionButton" onClick="Cancel('mrsIssuePickListEntry.asp')">
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
</body>
</html>
