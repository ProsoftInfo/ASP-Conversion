<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItemCycleCountHistoryPop.asp
	'Module Name				:	Inventory (Item Cycle Count)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 23,2013
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - UoM Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Selection.js"></SCRIPT>
</HEAD>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Cycle Count History
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
                                <td align="center" class="MiddlePack"></td>
                                <td width="100%">   
                                    <div style="height:415px;width:480px;">
                                        <table border="0" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
                                            <tr>
                                                <td class="ExcelHeaderCell" rowspan="2" align="center">S.No
                                                </td>
                                                <td class="ExcelHeaderCell" colspan="2" align="center">Current Stock</td>
                                                <td class="ExcelHeaderCell" colspan="2" align="center">Cycle Count</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center">Quantity</td>
                                                <td class="ExcelHeaderCell" align="center">Value</td>
                                                <td class="ExcelHeaderCell" align="center">Quantity</td>
                                                <td class="ExcelHeaderCell" align="center">Date</td>
                                            </tr>
                                            <%
                                                Dim rsObj
                                                Dim sQuery,sFinFrom,sFinTo,sArrPeriod,sArrTemp,sItemName,sOrgCode
                                                Dim iSNo,iCurrQty,iCurrVal,iCCQty,dCCDate,iCCVal,iItemCode,iClassCode
                                                
                                                set rsObj = Server.CreateObject("ADODB.Recordset")
                                                sArrperiod = split(session("FinPeriod"),":")
                                                sFinFrom = "01/04/"&sArrperiod(0)
                                                sFinTo = "31/03/"&sArrperiod(1)
                                                sArrTemp = split(Request("Info"),":")
                                                iItemCode = sArrTemp(0)
                                                iClassCode = sArrTemp(1)
                                                sOrgCode = sArrTemp(2)
                                                
                                                iSNo = 0
                                                sQuery = "Select CycleCountDate,IsNull(CycleCountStock,0),IsNull(CycleCountValue,0),IsNull(CurrentStock,0),IsNull(CurrentValue,0) from Inv_T_ItemCycleCountHistory where ClassificationCode = "& iClassCode &" and ItemCode = "& iItemCode &" and FinancialYearFrom = Convert(datetime,'"& sFinFrom &"',103) and FinancialYearTo=Convert(datetime,'"& sFinTo &"',103)"
                                                rsObj.open sQuery,con
                                                if not rsObj.eof then
                                                    do while not rsObj.eof
                                                        iSNo = iSNo + 1
                                                        dCCDate = rsObj(0)
                                                        iCCQty = rsObj(1)
                                                        iCCVal = rsObj(2)
                                                        iCurrQty = rsObj(3)
                                                        iCurrVal = rsObj(4)
                                                        %>
                                                            <tr>
                                                                <td class="ExcelSerial" align="center"><%=iSNo%></td>
                                                                <td class="ExcelDisplayCell" align="right"><%=FormatNumber(iCurrQty,3,0,0,0)%></td>
                                                                <td class="ExcelDisplayCell" align="right"><%=FormatNumber(iCurrVal,3,0,0,0)%></td>
                                                                <td class="ExcelDisplayCell" align="right"><%=FormatNumber(iCCQty,3,0,0,0)%></td>
                                                                <td class="ExcelDisplayCell" align="center"><%=dCCDate%></td>
                                                            </tr>
                                                        <%
                                                        rsObj.movenext
                                                    loop
                                                else
                                                    %>
                                                        <tr>
                                                            <td class="ExcelDisplayCell" colspan="5" align="center">No Cycle Counting Found</td>
                                                        </tr>
                                                    <%
                                                end if
                                                rsObj.close
                                            %>
                                        </table>
                                    </div>
                                </td>
                                <td align="center" class="MiddlePack"></td>
                            </tr>
                           	<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td valign="top" colspan="3">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Close" name="B2" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
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
<%
	' Function to populate UoM List
	Function populateUoMList()
		' Declaration of variables
		Dim dcrs,sUomDesc,sUomShDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT ORDER BY UOMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUoMCode = dcrs(0)
		set sUomDesc = dcrs(1)
		set sUomShDesc = dcrs(2)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUoMCode)&""">"&trim(sUomShDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
