<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	IssuePickedPackDetPop.asp
	'Module Name				:	Inventory (Picked Pack Details)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 28,2012
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
    Dim rsPick,rsTemp
    Dim sQuery,sOrgID,sOptionName,sItemDesc,sLotNo,sPackNo,sSerialNo
    Dim iPickNo,iItemCode,iClassCode,iAttID,iCtr,iPickQty
	
    set rsPick = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.RecordSet")
    
    sOrgID = session("organizationcode")
    
    Response.write "<font color=red>"
    
    iPickNo = Request("PickNo")
    iItemCode = Request("ItemCode")
    iClassCode = Request("ClassCode")
    iAttID = Request("AttID")
	
	sQuery = "Select ItemDescription from VWItem where ItemCode="& iItemCode
	rsTemp.open sQuery,con
	if not rsTemp.eof then
	    sItemDesc =  trim(rsTemp(0))
	end if
	rsTemp.close
	
	if trim(iAttID)<>"" then
        sQuery= "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iAttID
        rsTemp.open sQuery,con
        if not rsTemp.eof then
            sOptionName = trim(rsTemp(0))
        end if
        rsTemp.close
    end if
	
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function Fun_Close
window.close
End Function
</Script>
</head>
<body leftmargin="5" topmargin="2" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action>
<input type="hidden" name="hOrgID" value="<%=sOrgID%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="PageTitle" height="20">
			<p align="center">Picked Item Pack Details
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
											    <td class="FieldCellSub">Item Description&nbsp;</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sItemDesc%>&nbsp;</span></td>
											</tr>
											<tr>
											    <td class="FieldCellSub">Attribute Name</td>
											    <td class="FieldCell"><span class="DataOnly"><%=sOptionName%>&nbsp;</span></td>
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
												    <div class="frmBody" id="frm3" style="width: 100%; height:235">
													    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
														    <tr>
															    <td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
															    <td class="ExcelHeaderCell" align="center">Lot Number</td>
															    <td class="ExcelHeaderCell" align="center">Packing Number</td>
															    <td class="ExcelHeaderCell" align="center" width="20%">Picked Quantity</td>
														    </tr>
													    <%
													        Response.Write "<font color=red>"
													        
													        sQuery = "Select PickNumber,QuantityPicked,isNull(SerialNo,0) from INV_T_IssuePickDetails where PickNumber = "& iPickNo &" and ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode 
													        if trim(iAttID)<>"" and not isNull(iAttID) then
													            sQuery = sQuery &" and ItemAttributes = "& iAttID
													        end if 
													        
													        'Response.write "<textarea>"&sQuery&"</textarea>"
													        rsPick.open sQuery,con
													        if not rsPick.eof then
													            do while not rsPick.eof
													                iCtr = iCtr + 1
													                iPickQty = rsPick(1)
													                sSerialNo =trim(rsPick(2))
													                
													                if trim(sSerialNo)<>"0" then
													                    sQuery = "Select isNull(LotNumber,0),isNull(PackingNumber,0) from INV_T_LocationLot where SerialNumber = "&sSerialNo
													                    rsTemp.open sQuery,con
													                    if not rsTemp.eof then
													                        sLotNo = rsTemp(0)
													                        sPackNo = rsTemp(1)
													                    end if
													                    rsTemp.close
													                else
													                        sLotNo = 0
													                        sPackNo = 0
													                end if
													                if trim(sLotNo)="" or isNull(sLotNo) or sLotNo="0" then sLotNo = "N/A"
													                if trim(sPackNo)="" or isNull(sPackNo) or sPackNo="0" then sPackNo = "N/A"
													                
													                %>
													                    <tr>
													                        <td class="ExcelDisplayCell" align="center"><%=iCtr%></td>
													                        <td class="ExcelDisplayCell"><%=sLotNo%></td>
													                        <td class="ExcelDisplayCell" align="right"><%=sPackNo%></td>
													                        <td class="ExcelDisplayCell" align="right"><%=FormatNumber(iPickQty,2)%></td>
													                    </tr>
													                <%
													                rsPick.MoveNext
													            loop
													        end if
													        rsPick.close
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
 												<input type="button" value="Close" name="cancel" class="ActionButton" onClick="Fun_Close()">
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