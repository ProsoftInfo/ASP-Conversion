<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ITEMCODESETUP.asp
	'Module Name				:	Inventory (Item Code Setup)
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 04,2013
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	ItemCodeSetupInsert.asp
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Display</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
	function CheckSubmit()
	{
		if ((document.formname.ChkItemCode.checked) &&  (document.formname.txtItemCode.value ==""))
		{
			alert ("Enter Item code display order");
			document.formname.txtItemCode.focus();
			return false;
		}
		else if ((document.formname.ChkItemName.checked) &&  (document.formname.txtItemName.value ==""))
		{
			alert ("Enter Item Name display order");
			document.formname.txtItemName.focus();
			return false;
		}
		else if ((document.formname.ChkDrawing.checked) &&  (document.formname.txtDrawing.value ==""))
		{
			alert ("Enter Drawing Version No display order");
			document.formname.txtDrawing.focus();
			return false;
		}
		else if ((document.formname.ChkCatalog.checked) &&  (document.formname.txtCatalog.value ==""))
		{
			alert ("Enter Catalogue No display order");
			document.formname.txtCatalog.focus();
			return false;
		}
		else if (!(document.formname.ChkItemCode.checked) &&  !(document.formname.txtItemCode.value ==""))
		{
			alert ("Select Item code");
			document.formname.ChkItemCode.focus();
			return false;
		}
		else if (!(document.formname.ChkItemName.checked) &&  !(document.formname.txtItemName.value ==""))
		{
			alert ("Select Item Name");
			document.formname.ChkItemName.focus();
			return false;
		}
		else if (!(document.formname.ChkDrawing.checked) &&  !(document.formname.txtDrawing.value ==""))
		{
			alert ("Select Drawing Version No");
			document.formname.ChkDrawing.focus();
			return false;
		}
		else if (!(document.formname.ChkCatalog.checked) &&  !(document.formname.txtCatalog.value ==""))
		{
			alert ("Select Catalogue No ");
			document.formname.ChkCatalog.focus();
			return false;
		}
		else if (!((document.formname.ChkItemCode.checked) ||  (document.formname.ChkItemName.checked) || (document.formname.ChkDrawing.checked) || (document.formname.ChkCatalog.checked)))
		{
			alert ("Select Item Code ");
			document.formname.ChkItemCode.focus();
			return false;
		}
		else
			document.formname.submit();
	}

	function ClearText()
	{
		document.formname.txtItemCode.value = ""
		document.formname.txtItemName.value = ""
		document.formname.txtDrawing.value = ""
		document.formname.txtCatalog.value = ""
	}	
</Script>
<Script>
	function Check()
	{
		var form = document.formname;
		var flag = false;

		if (form.hItemCode.value != "0") {
			form.ChkItemCode.checked = true;
			form.txtItemCode.value = form.hItemCode.value;
			flag = true;
		}
		if (form.hItemName.value != "0") {
			form.ChkItemName.checked = true;
			form.txtItemName.value = form.hItemName.value;
			flag = true;
		}
		if (form.hDrawingNo.value != "0") {
			form.ChkDrawing.checked = true;
			form.txtDrawing.value = form.hDrawingNo.value;
			flag = true;
		}
		if (form.hCatalogNo.value != "0") {
			form.ChkCatalog.checked = true;
			form.txtCatalog.value = form.hCatalogNo.value;
			flag = true;
		}

		form.B2.value = flag ? "Amend" : "Save";
	}
</Script>

</HEAD>
<%
	dim dcrs1, iItemCode, iItemName, iDrawingNo, iCatalogNo
	
	iItemCode = 0
	iItemName = 0 
	iDrawingNo = 0
	iCatalogNo = 0
	
	set dcrs1 = server.CreateObject ("ADODB.Recordset")
	
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISPLAYNAME, DISPLAYORDER FROM INV_M_ITEMDISPLAY ORDER BY DISPLAYORDER"
		.ActiveConnection = con
		.Open 
	end with
	set dcrs1.ActiveConnection = Nothing
		
	do while not dcrs1.EOF 
		if dcrs1(0)="INO" then
			iItemCode = dcrs1(1)
		elseif dcrs1(0)="INA" then
			iItemName = dcrs1(1)
		elseif dcrs1(0)="DNO" then
			iDrawingNo = dcrs1(1)
		elseif dcrs1(0)="CNO" then
			iCatalogNo = dcrs1(1)
		end if
		dcrs1.MoveNext 
	loop
	dcrs1.Close 
	
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Check()">
<form method="POST" name="formname" action="ItemCodeSetupInsert.asp">
<input type=hidden name="hItemCode" value="<%=iItemCode%>">
<input type=hidden name="hItemName" value="<%=iItemName%>">
<input type=hidden name="hDrawingNo" value="<%=iDrawingNo%>">
<input type=hidden name="hCatalogNo" value="<%=iCatalogNo%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Code Setup</p>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="BodyTable">
										<tr>
											<td width="100%" align="center">
													<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width=50%>
														<tr>
															<td class="ExcelHeaderCell" align="Center" width="10">Select</td>
															<td class="ExcelHeaderCell" align="Center" width="300">Display By</td>
														    <td class="ExcelHeaderCell" align="Center" width="50">Display Order</td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkItemCode" class="Formelem" onClick="javascript:ClearText()"></td>
															<td class='ExcelDisplayCell'> Item Code</td>
															<td class='ExcelInputCell'><input type="text" name="txtItemCode" class="Formelem" size="10" onKeyPress="javascript:DoKeyPress('N',1,0)"></td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkItemName" class="Formelem" onClick="javascript:ClearText()"></td>
															<td class='ExcelDisplayCell'> Item Name</td>
															<td class='ExcelInputCell'><input type="text" name="txtItemName" class="Formelem" size="10" onKeyPress="javascript:DoKeyPress('N',1,0)"></td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkDrawing" class="Formelem" onClick="javascript:ClearText()"></td>
															<td class='ExcelDisplayCell'> Drawing Version No.</td>
															<td class='ExcelInputCell'><input type="text" name="txtDrawing" class="Formelem" size="10" onKeyPress="javascript:DoKeyPress('N',1,0)"></td>
														</tr>
														<tr>
															<td class='ExcelDisplayCell' align="center"><input type="Checkbox" name="ChkCatalog" class="Formelem" onClick="javascript:ClearText()"></td>
															<td class='ExcelDisplayCell'> Catalouge No.</td>
															<td class='ExcelInputCell'><input type="text" name="txtCatalog" class="Formelem" size="10" onKeyPress="javascript:DoKeyPress('N',1,0)"></td>
														</tr>
													</table>
											</td>
										</tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Save" name="B2" class="ActionButton" onClick="javascript:CheckSubmit()">
																<input type="reset" value="Reset" name="B1" class="ActionButton">
														</td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
									</table>
								</td>
								<td align="center">
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

