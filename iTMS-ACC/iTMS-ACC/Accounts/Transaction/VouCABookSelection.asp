<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouCashBookSelection.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 21, 2002
	'Modified On				:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<!-- XML Data Island -->
<XML ID="UnitBookData">
<Book/>
</XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../scripts/VouSelection.js"></SCRIPT>
<SCRIPT language="vbscript">



Function DisplayBook(objUnit)
dim iUnitNo,arrTemp
dim Root
	document.formname.selBook.options.length = 1

	if objUnit.selectedIndex <> "0" then
		iUnitNo= objUnit(objUnit.selectedIndex).value

		set objhttp = CreateObject("MSXML2.XMLHTTP")

		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=01&orgID=" & iUnitNo , false
		objhttp.send
		'alert objhttp.responsetext

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement

			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if
	end if
end Function

Function VouCreate
	if validate then
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes

			if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
				document.formname.hBookAccHead.value=HeaderNode.Attributes.Item(2).nodeValue
				document.formname.hBookOtherUnit.value=HeaderNode.Attributes.Item(3).nodeValue

				if HeaderNode.Attributes.Item(4).nodeValue="C" and document.formname.selVouType.value="C" then
					MsgBox ("Book balance is in Credit cannot make Payment")
					document.formname.selVouType.focus
					VouCreate= false
					exit function
				end if
			end if
		next

		document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
		document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
		document.formname.action="VouCAEntry.asp"
		document.formname.submit()
	end if
End function

Function VouAmend
dim iBookNo,sOrgId,sVouNo,iBookId,sTrans
	iBookNo=document.formname.selBook.value
	sOrgId=document.formname.selUnitId.value
	sTrans=document.formname.selVouType.value
	sVouNo=document.formname.txtVouNo.value

	iBookId="01"
	if sTrans="C" then
		sTrans="CAP"
	else
		sTrans="CAR"
	end if
	if validate then
		IF document.formname.txtVouNo.value = "" Then
			MsgBox "Select Voucher Number "
			exit Function
		End IF
		iTranNo = document.formname.hTransNo.value

		'set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","XMLVouNoValidate.asp?orgID=" & sOrgId&"&BookId="&iBookId&"&BookNo="&iBookNo&"&VouNo="&sVouNo&"&Trans="&sTrans, false
		'objhttp.send

		'if  cint (objhttp.responseText)>0 then
			document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
			document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
			document.formname.hTransNo.value=iTranNo

			Set Root = UnitBookData.documentElement
			For Each HeaderNode In Root.childNodes
				if  HeaderNode.Attributes.Item(0).nodeValue=document.formname.selBook.value then
					document.formname.hBookAccHead.value=HeaderNode.Attributes.Item(2).nodeValue
					document.formname.hBookOtherUnit.value=HeaderNode.Attributes.Item(3).nodeValue
				end if
			next
			document.formname.action="VouCAAmdEntry.asp"
			document.formname.submit()
		'else
			'MsgBox ("Not a Valid Voucher Number ")
			'document.formname.txtVouNo.select
		'end if
	end if

End function

Function VouDel
dim iBookNo,sOrgId,sVouNo,iBookId,sTrans
	iBookNo=document.formname.selBook.value
	sOrgId=document.formname.selUnitId.value
	sTrans=document.formname.selVouType.value
	sVouNo=document.formname.txtVouNo.value

	iBookId="01"
	if sTrans="C" then
		sTrans="CAP"
	else
		sTrans="CAR"
	end if
	if validate then
		IF document.formname.txtVouNo.value = "" Then
			MsgBox "Select Voucher Number "
			exit Function
		End IF
		iTranNo = document.formname.hTransNo.value
		'set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","XMLVouNoValidate.asp?orgID=" & sOrgId&"&BookId="&iBookId&"&BookNo="&iBookNo&"&VouNo="&sVouNo&"&Trans="&sTrans, false
		'objhttp.send

		'if cint (objhttp.responseText)>0 then
			document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
			document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
			document.formname.hTransNo.value=iTranNo
			document.formname.action="VouCADelDisplay.asp"
			document.formname.submit()
		'else
			'MsgBox ("Not a Valid Voucher Number ")
			'document.formname.txtVouNo.select
		'end if

	end if
End function

Function VouView
dim iBookNo,sOrgId,sVouNo,iBookId,sTrans
	iBookNo=document.formname.selBook.value
	sOrgId=document.formname.selUnitId.value
	sTrans=document.formname.selVouType.value
	sVouNo=document.formname.txtVouNo.value

	iBookId="01"
	if sTrans="C" then
		sTrans="CAP"
	else
		sTrans="CAR"
	end if

	if validate then
		IF document.formname.txtVouNo.value = "" Then
			MsgBox "Select Voucher Number "
			exit Function
		End IF
		iTranNo = document.formname.hTransNo.value

		'set objhttp = CreateObject("MSXML2.XMLHTTP")
		'objhttp.Open "GET","XMLVouNoValidate.asp?orgID=" & sOrgId&"&BookId="&iBookId&"&BookNo="&iBookNo&"&VouNo="&sVouNo&"&Trans="&sTrans, false
		'objhttp.send


		'if  cint (objhttp.responseText)>0 then
			document.formname.hBookName.value=document.formname.selBook.options(document.formname.selBook.selectedIndex).text
			document.formname.horgName.value=document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).text
			'document.formname.hTransNo.value=objhttp.responseText
			document.formname.hTransNo.value=iTranNo
			document.formname.action="VouCAView.asp"
			document.formname.submit()
		'else
		'	MsgBox ("Not a Valid Voucher Number ")
		'	document.formname.txtVouNo.select
		'end if
	end if
End function
function validate()
	if document.formname.selUnitId.selectedIndex<1 then
		MsgBox ("Select Unit")
		validate= false
		exit function
	end if
	if document.formname.selBook.selectedIndex<1 then
		MsgBox ("Select Cash Book")
		validate= false
		exit function
	end if
	if document.formname.selVouType.selectedIndex<1 then
		MsgBox ("Select Voucher type")
		validate=false
		exit function
	end if
	validate=true
End function



</script>



</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<input type="hidden" name="hBookName" value="">
<input type="hidden" name="hBookAccHead" value="">
<input type="hidden" name="hBookOtherUnit" value="">

<input type="hidden" name="horgName" value="">
<input type="hidden" name="hTransNo" value="">
<input type="hidden" name="hAmendTy" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Cash Voucher
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">

								  		<tr><td align="center">Voucher</td>
								  	</tr>
								  </table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                            <td class="FieldCell" width="108">Organization </td>
                            <td class="FieldCell">
                             <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationListDB%>
                              </select>
                                            </td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Book</td>
                            <td class="FieldCell">
                            <select size="1" name="selBook" class="FormElem">
                        <option value="S">Select Book</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Voucher Type</td>
                            <td class="FieldCell">
                            <select size="1" name="selVouType" class="FormElem">
							<option value="S">Select Voucher Type</option>
							<option value="C">Payment</option>
							<option value="D">Receipt</option>
                            </select></td>
                                </tr>
                                <tr>
                            <td class="FieldCell" width="108">Voucher Number</td>
                            <td class="FieldCell">
                            <input type="text" name="txtVouNo" size="20" class="FormElem" readonly>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                           <a href="javascript:popVoucherNo('C','01','CA',document.formname.selUnitId.value,document.formname.selVouType.value,document.formname.selBook.value)">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Vouchers Created Not Accounted"></a>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:popVoucherNo('A','01','CA',document.formname.selUnitId.value,document.formname.selVouType.value,document.formname.selBook.value)">
                           <img border="0" src="../../assets/images/iTMS Icons/Details.gif" alt="Accounted Vouchers"></a>
                            </td>

                                </tr>
                                    </table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Create" onClick="VouCreate()" name="btnCreate" class="ActionButton" >
                                                                <input type="button" value="View" name="B7" onClick="VouView()" class="ActionButton">
                                                                <input type="button" value="Amendment" name="btnAmend" onClick="VouAmend()"  class="ActionButtonX">
                                                                <input type="button" value="Delete" name="btnDel" onClick="VouDel()" class="ActionButton">
                                                                <input type="reset" value="Reset" name="B10" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="10" class="BottomPack" colspan="3">
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