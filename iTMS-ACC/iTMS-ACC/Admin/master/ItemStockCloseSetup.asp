<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TransferClosingEntry.asp
	'Module Name				:	Transfer Closing Values
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 12, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	TransferClosingDetailsEntry.asp
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Transfer Closing</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<xml id="SubCategory"><Root></Root></xml>
<xml id="Classification"><Root></Root></xml>
<xml id="OutDataXML"><Root></Root></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ItemStockCloseSetupCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
ITMSItemStockCloseSetupCompat.install();
</SCRIPT>
</HEAD>
<BODY leftMargin=20 topMargin=15 onload="CheckSetup();LoadXML();DisplayTable()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hRow" value="0">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

	<tr>
		<td align="center" class=PageTitle height="20">
			<p align="center">
			    Item Stock Closing Setup
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
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%" align="left">
								    <div style="height:450">
                                        <table BORDER="0" CELLSPACING="1" CELLPADDING="0" width="100%">
                                            <tr>
                                                <td class="FieldCell">
                                                    <input type="checkbox" name="chkCategory" onClick="CheckSetup()">Category
                                                </td>
                                                <td class="FieldCell">
                                                    <input type="checkbox" name="chkSubCategory" onClick="CheckSetup()">Sub Category
                                                </td>
                                                <td class="FieldCell">
                                                    <input type="checkbox" name="chkClassification" onClick="CheckSetup()">Classification
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="FieldCell">
                                                    <select name="selCategory" Size="5" class="FormElem" onchange="populateSubCategory()" onclick="populateSubCategory()">
                                                        <%
                                                            populateCategory()
                                                        %>
                                                    </select>
                                                </td>
                                                <td class="FieldCell">
                                                    <select name="selSubCategory" Size="5" class="FormElem" onchange="populateClassification()" onclick="populateClassification()" multiple>
                                                    </select>
                                                </td>
                                                <td class="FieldCell">
                                                    <select name="selClassification" Size="5" class="FormElem" multiple>
                                                    </select>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="FieldCell">
                                                    <input type="Button" name="btnCategory" class="ActionButton" value="Add" onclick="AddCategorySetup()">
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="3">
                                                    <table width="100%" id="tblSetup" class="ExcelTable">
                                                        <tr>
                                                            <td class="ExcelHeaderCell" align="center">S.No</td>
                                                            <td class="ExcelHeaderCell" align="center">
                                                                <img src="../../assets/images/iTMS%20icons/DeleteIcon.gif" onclick="DeleteSetUp()">
                                                            </td>
                                                            <td class="ExcelHeaderCell" align="center">Category</td>
                                                            <td class="ExcelHeaderCell" align="center">Sub Category</td>
                                                            <td class="ExcelHeaderCell" align="center">Classification</td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
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
                                                    <input type="button" name="btnSetup" value=" Save " onclick="SetupCategory()" class="ActionButtonX">
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

<%
Function populateCategory()
    Dim rsObj,sQuery
    set rsObj = Server.CreateObject("ADODB.Recordset")
    sQuery = "Select CategoryCode,CategoryName from Inv_M_ClassificationCategory"
    rsObj.Open sQuery,con
    if not rsObj.EOF then
        do while not rsObj.EOF 
            Response.Write "<option value="& Trim(rsObj(0)) &">"&trim(rsObj(1))&"</option>"
            rsObj.MoveNext 
        loop
    end if 
    rsObj.Close 
End Function
%>