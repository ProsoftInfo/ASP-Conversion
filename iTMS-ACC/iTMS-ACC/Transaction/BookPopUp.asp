<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BookPopUp.asp
	'Module Name				:	ACCOUNTS ()
	'Author Name				:
	'Modified By				:	S.Maheswari
	'Created On					:	Sep 16 2008
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
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%

'XML DOM Variables
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS -
	  <%

			Response.Write "Select Book"

	%>
</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML ID="UnitBookData"><Book/></XML>
<script language="vbscript">
window.ReturnValue = "0--0"
Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
	document.formname.selBook.options.length = 0
	iUnitNo = document.formname.hUnitId.value
	'alert iUnitNo
	'if objUnit.selectedIndex <> "0" then
		'iUnitNo= objUnit(objUnit.selectedIndex).value
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		if trim(document.formname.hVouType.Value) = "GJ" then
			objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=08&orgID=" & iUnitNo , false
		elseif trim(document.formname.hVouType.Value) = "CR"  then
			objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=07&orgID=" & iUnitNo , false
		elseif trim(document.formname.hVouType.Value) = "DR" then
			objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=06&orgID=" & iUnitNo , false
		end if
		objhttp.send

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement
			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if

end Function
Function Win_UnLoad()
     if document.formname.selBook.selectedIndex = - 1 then
        alert("Select Book")
        exit function
     end if
	window.ReturnValue = document.formname.selBook.options(document.formname.selBook.selectedIndex).value &"--"& document.formname.selBook.options(document.formname.selBook.selectedIndex).text
	window.close()
End Function
Function window_onunload()
	window.close()
End Function
</script>
<%
Dim sUnit,sVouType
sUnit = Request("Unit")
sVouType = Request("VouType")

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload= "DisplayBook()" >
<form method="POST" name="formname" action="">
<Input type="hidden" name="hUnitId" value="<%=sUnit%>">
<Input type="hidden" name="hVouType" value="<%=sVouType%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Book
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center" width="5">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly" width="100%">

											 <tr>
											    <td class="FieldCellSub" width="168">Book</td>
											    <td class="FieldCell">
											    <select size="5" name="selBook" class="FormElem">
												<!--option value="S">Select Book</option-->
											    </select></td>
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
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
 													<input type="button" value="Done" name="B3" class="ActionButton" onclick="Win_UnLoad()">

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
