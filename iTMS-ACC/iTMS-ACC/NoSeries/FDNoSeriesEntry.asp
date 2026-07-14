<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SalesNoSeriesEntry.asp
	'Module Name				:	Sales (Master Creation)
	'Author Name				:	Subbiah
	'Created On					:	August 07,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	SalesNoSeriesInsert.asp
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
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="SeriesNoData" data-src="../NoSeries/xmldata/SeriesNumberDetail.xml"></script>
<SCRIPT LANGUAGE=javascript SRC="../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/NoSeriesEntryCompat.js"></SCRIPT>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" onLoad="popSeriesNo()" MARGINWIDTH="0">

<form method="POST" name="formname" action="FDNoSeriesInsert.asp">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hActivityName" value="">
<input type=hidden name="hOptval" value="N">
<input type=hidden name="hAgentName" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">No. Series Allocation</p>
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
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<!--<tr>
											<td class=FieldCell> Select Unit</td>
											<td class="FieldCellSub">
												<select size="1" name="selUnit" class="FormElem">
													<option value="0">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														'populateOrganizationList
													%>
												</select>
											</td>
										</tr>-->
										<tr>
											<td class="FieldCell">Activity</td>
											<td class="FieldCellSub">
											    <select size="1" name="selActType" class="FormElem" onchange ="Activity_change(this)">
													<option value="0">Select</option>
													<option value="FDR">Deposit Receipt No</option>
													<option value="FLN">Folio No</option>
													<option value="LON">Loan Sanction No</option>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Deposit Type</td>
											<td class="FieldCell">
												<INPUT type="radio" id=radio1 name="optInv" value = "F" onclick = "Optclick(this)"> Fixed
												<INPUT type="radio" id=radio1 name="optInv" value = "C" onclick = "Optclick(this)"> Cumlative

											</td>
										</tr>



										<tr>
											<td class=FieldCell> Select No Series</td>
											<td class="FieldCellSub">
												<select size="1" name="selNoSeries" class="FormElem" onChange="DisplayTable()">
													<OPTION value="0">Select</option>
												</select>
                                            </td>
										</tr>
									</table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td align="center" valign="top">
                                         <table id="tblBook" border="0" cellspacing="1" class="ExcelTable" >
                                            <tr>
											<td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.</td>
											<td class="ExcelHeaderCell" align="center" width="75">Period</td>
											<td class="ExcelHeaderCell" align="center" width="50">Start No</td>
											<td class="ExcelHeaderCell" align="center" width="100">Prefix</td>
											<td class="ExcelHeaderCell" align="center" width="100">Suffix</td>
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
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															<input type="button" value="Submit" name="B4" class="ActionButton" onClick="validateForm()" >
															<input type="Reset" value="Cancel" name="B2" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../assets/images/clearpixel.gif" width="5" height="5">
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
	' Function to populate the Item Type list
	Function populateSaleType()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT InvoiceType,InvoiceTypeName FROM Sal_M_InvoiceTypes  ORDER BY InvoiceType"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function



	' Function to populate the Item Type list
	Function populateAgent()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT AgentCode,AgentName FROM APP_M_AgentMaster  ORDER BY AgentName"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>

