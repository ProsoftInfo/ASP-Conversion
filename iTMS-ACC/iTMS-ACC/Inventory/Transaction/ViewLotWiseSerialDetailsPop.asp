<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ViewLotWiseSerialDetailsPop.asp
	'Module Name				:	Inventory 
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	April 23,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None

%>
<%
	Dim rsTemp,objrs
	Dim sOrgCode,sQuery,sItemCode,sLotQuantity,sTotLotQuantity,sLotNumber,iSNo
	sOrgCode = Session("organizationcode")
	sItemCode = Request.QueryString("ItemCode")
	sLotNumber = Request.QueryString("LotNumber")
	set rsTemp = Server.CreateObject("ADODB.Recordset")
	set objrs = Server.CreateObject("ADODB.Recordset")
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Lot Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="RefType"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="RefData"><Root Done="N"/></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root></Root></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=20 topMargin=10>
<form method="POST" name="formname">
<input type=hidden name="hUnit" value="<%=sOrgCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" >
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Lot Details
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
					<td width="10px"></td>
					<TD class=TabBodywithtopline>

						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							
							<tr>
								<td align="left" colspan=3>
								    <div style="width:390px;height:290px">
								        <table class=ExcelTable id="tblLot" border=0 width=100%>
								            <tr>
								                <td class="ExcelHeaderCell" align=center>S.No</td>
								                <td class="ExcelHeaderCell" align=center>Lot No</td>
								                <td class="ExcelHeaderCell" align=center>Quantity</td>
								            </tr>
								            <%
								                sQuery = "Select isNull(LotNumber,''),isNull(Sum(LotQuantityNett),0)-isNull(Sum(QuantityIssued),0) from INV_T_LocationLot "&_
								                " where Itemcode = "& sItemCode &" and LotQuantityNett-QuantityIssued > 0  Group By LotNumber"
								                objrs.Open sQuery,con
								                if not objrs.EOF then
								                    iSNo  = 0
								                    sLotNumber = 0
								                    sLotQuantity = 0
								                    sTotLotQuantity = 0
								                    do while not objrs.EOF 
								                        sLotNumber = objrs(0)
								                        sLotQuantity = objrs(1)
								                        sTotLotQuantity = cdbl(sTotLotQuantity)+cdbl(sLotQuantity)
    								                    
    								                        if Trim(sLotNumber)="" or IsNull(sLotNumber) or sLotNumber="0" then  sLotNumber = "N/A"
								                        
								                            iSNo = iSNo + 1
								                            %>
								                                <tr>
								                                    <td class="ExcelSerial" align=center><%=iSNo%></td>
								                                    <td class="ExcelDisplayCell" align=center><%=sLotNumber%></td>
								                                    <td class="ExcelDisplayCell" align=center><%=sLotQuantity%></td>
								                                </tr>
								                            <%
								                        objrs.MoveNext     
								                    loop
								                end if
								                objrs.Close 
								            %>
								        </table>
								    </div>
								</td>
							</tr>
							
							<tr>
								<td align="center" colspan="3" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="5" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="5" class="ActionCell" >
									<input type=button name="btnClose" value="Close" class="ActionButtonX" onclick="window.close()">
								</td>
							</tr>


                        </table>
					</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
