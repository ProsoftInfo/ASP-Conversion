<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BooksCreationEntry.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 04, 2002
	'Modified By                :   Ragavendran R
	'Modified On				:   Jan 19,2011
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
dim objRs,objRs1,objFs,sQuery
Dim oDOM,Root,newElem,newElem1,nodUnit
dim sUnitID,sUnitLName,sUnitSName,sBookID,sBookName

Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objfs = CreateObject("Scripting.FileSystemObject")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

sUnitID = Session("organizationcode")
sUnitSName = Session("OrgShortName")

'oDOM.Save server.MapPath("../xmldata/UnitBookDetails.xml")

'if not objFs.FileExists(server.MapPath("../../NoSeries/xmldata/SeriesNumberDetail.xml")) then
%>
<!--<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT>
<!--
	function msgbox(strr)
	{
			alert(strr);
			window.location.href = "../AccountsHome.asp";
	}
//-->
<!--</SCRIPT>
<BODY onLoad = "msgbox('Number Series Not Defined')">
</BODY>
<HTML>-->

<%
'Response.End
'end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="SeriesNoData" data-src="../../NoSeries/xmldata/SeriesNumberDetail.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="UnitBook"><Root /></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "booksCreationEntry" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" onLoad="CheckNumberSerious();popSeriesNo();popUnitBooks()" MARGINWIDTH="0">

<form method="POST" name="formname" action="BooksCreationUpdate.asp">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hUnitID" value="<%=sUnitID%>">
<input type=hidden name="hUnitName" value="<%=sUnitSName %>" >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Day Book Creation</p>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							  <tr>
								<td align="center"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td width="100%" align="left">
									<table border="0" cellspacing="0"  cellpadding="0" class="ToolBarTable">
										<tr>
										<td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
					<a href="#" onclick="popDayBookList(); return false;"><span style="cursor: pointer" Title="View Contra Details" >
              						      <p align="center"><font face="Wingdings" color="#000000" size="5">4</font>
                                        </span></a>
					                    </td>

										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0" width="100%">
										<!--<tr>
											<td class=FieldCell width="200"> Select Unit</td>
											<td><select size="1" name="selUnitId" class="FormElem">
											<OPTION value="0">Select a Unit</option>
											<%populateOrganizationList%>
                                            </select></td>
										</tr>-->
                                        <tr>
                                        	<td align="center" class="MiddlePack" colspan="2">
                                        	</td>
                                        </tr>
										<tr>
											<td class=FieldCell width="200"> Select Day Book</td>
											<td>
												<select size="1" name="selDayBook" onChange="setPayRec()" class="FormElem" >
													<OPTION value="0">Select a Day Book</option>
													<%
													    sQuery = "select BookCode,BookName from Acc_M_DayBooks"
													    objrs.open sQuery,con
													    if not objrs.eof then
													        do while not objrs.eof
													            Response.Write "<option value="&trim(objrs(0))&">"& trim(objrs(1)) &"</option>"
													            objrs.movenext
													        loop
													    end if
													    objrs.close
													%>
												</select>
                                            </td>
										</tr>
                                        <tr>
                                        	<td align="center" class="MiddlePack" colspan="2">
                                        	</td>
                                        </tr>
										<tr>
											<td class=FieldCell width="200">Allow Other&nbsp;Units Transaction
											</td>
											<td>
												<table border="0" cellpadding="0" cellspacing="0">
													<tr>
														<td width="20"><input type="radio" value="1" name="optEligible" checked class="formelem"></td>
														<td class="FieldCell" width="30">Yes </td>
														<td width="20"><input type="radio" value="0" name="optEligible" class="formelem"></td>
														<td class="FieldCell">No</td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
                                        	<td align="center" class="MiddlePack" colspan="2">
                                        	</td>
                                        </tr>
										<tr>
											<td class=FieldCell width="200"> Book Name</td>
											<td><input type="text" class="Formelem" maxlength="50" name="txtName" size="45"></td>
										</tr>
                                        <tr>
                                        	<td align="center" class="MiddlePack" colspan="2">
                                        	</td>
                                        </tr>
                                        <tr>
											<td class=FieldCell width="200"> Separate Payment / Receipt No&nbsp;</td>
											<td>
											<select size="1" name="selPayRecNo" class="FormElem" onChange="DisplayBook()">
											<OPTION value="Y">Yes</option>
											<OPTION value="N">No</option>
                                            </select></td>
                                        </tr>
                                        <tr>
                                        	<td align="center" class="MiddlePack" colspan="2">
                                        	</td>
										<tr>
											<td class=FieldCell width="200"> Select No Series</td>
											<td><select size="1" name="selNoSeries" class="FormElem" onChange="DisplayBook()">
											<OPTION value="0">Select Number Series</option>
                                            </select></td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															<input type="button" value="Save" name="B4" class="ActionButton" onClick="validateForm()" >
															<input type="button" value="Cancel" name="B2" onClick="Cancel('DayBookGrid.asp')"  class="ActionButton" >
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
