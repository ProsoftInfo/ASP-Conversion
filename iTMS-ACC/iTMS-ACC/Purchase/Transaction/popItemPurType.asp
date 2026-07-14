<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	popItemPurType.asp
	'Module Name				:	Purchase (Transactions-Invoice)
	'Author Name				:
	'Created On					:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	invPurInvoiceHeaderEntry.asp
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
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS-Invoice Form Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<%
dim oDOM,oNodRoot,objFSO
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set objFSO = Server.CreateObject("Scripting.FileSystemObject")

if objFSO.FileExists(server.MapPath("../temp/transaction/InvItemValue_PUR_"&Session.SessionID&".xml")) then
    oDOM.load server.MapPath("../temp/transaction/InvItemValue_PUR_"&Session.SessionID&".xml")
else
    oDOM.load server.MapPath("../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml")
end if
set oNodRoot=oDOM.documentElement
%>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/popItemPurType.js"></SCRIPT>
<script type="application/xml" id="TaxFormData" data-itms-xml-island="1"><Root/></script>
<script type="text/plain" data-itms-legacy-client-script="1">
Dim InvoiceDet
Set InvoiceDet = window.dialogArguments
'----------------------------------------------------
function Submit_Clk()

dim root,ItemNode,i,ItemCode,ClassCode,purType,DeleteNode,TaxNode,RtTax,ndEntry,objhttp,objhttp1,sForUnit
Set objhttp = CreateObject("Microsoft.XMLHTTP")
Set objhttp1 = CreateObject("Microsoft.XMLHTTP")

set root =InvoiceDet.DocumentElement

'To get Org Id
Set ItemNode=Root.Selectnodes("//InvoiceHeader/Header")
if ItemNode.Length>0 then
	sForUnit=ItemNode.Item(i).Attributes.getNamedItem("OrgID").value
end if

'To delete Existing Taxdetails Node
Set DeleteNode = Root.Selectnodes("//TaxDetails")
for i=0 to (DeleteNode.Length-1)
	root.RemoveChild(DeleteNode.Item(i))
next

'To set purchase type for each Item and add TaxDetails Node
if Root.hasChildNodes then
	Set ItemNode=Root.Selectnodes("//ItemDetails/Item")
	for i = 0 to ItemNode.Length - 1
		ItemCode=ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
		ClassCode=ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
		EntNo=ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value
		purType=eval("document.formname.cmbPurType" +ClassCode + "Z" + ItemCode + "Z" + EntNo).value
		 'msgbox "Purtype="&purType
		ItemNode.Item(i).Attributes.getNamedItem("PurchaseType").value= purType
		'Set TaxNode=Root.Selectnodes("//TaxDetails[@PurchaseType="& purType &"]")
		'alert(TaxNode.Length)
		'if TaxNode.Length=0 then
			objhttp.Open "GET","XMLGetTaxDetails.asp?PurType="&purType&"&ForUnit="&sForUnit,false
			objhttp.send
			'alert("objhttp.responsetext="&objhttp.responsetext)
			'alert("objhttp.responseXML.xml"+objhttp.responseXML.xml)
			If objhttp.responsexml.xml <> "" then
				TaxFormData.loadXML objhttp.responseXML.xml
				Set RtTax = TaxFormData.documentElement
				'alert("RtTax="&RtTax.xml)
				If RtTax.hasChildNodes Then
					Set Root=InvoiceDet.documentElement
					For Each ndEntry in RtTax.childNodes
						Root.appendchild ndEntry
					next
				end if	'If RtTax.hasChildNodes
			end if	'If objhttp.responsexml.xml
		'end if	'if TaxNode.Length
	next
end if
'alert("Chk="&Root.xml)
'set objhttp1 = CreateObject("Microsoft.XMLHTTP")
'objhttp1.Open "POST","XMLSavePur.asp?Mod=PUR&Name=InvItemValue1", false
'objhttp1.send InvoiceDet.XMLDocument
'alert "saved"
window.close
end function

'-------------------------------------------------------------------------------------------
Function window_onunload()
	Set window.returnvalue= InvoiceDet.documentElement
End Function
'-------------------------------------------------------------------------------------------

</script>


</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
          Purchase Item Tax
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
<TD class="TabBodyWithTopLine">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
                <tr>
					<td align="center" colspan="3" class="MiddlePack" height="7" width="50%">
					</td>
                </tr>
                <tr>
					<td align="center" width="5" class="ClearPixel" height="2">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
					<td align="center" class="ClearPixel" width="6" height="2">
                        <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
                </tr>
                <tr>
					<td align="center" colspan="3" class="MiddlePack" height="7" width="50%">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
                </tr>
                <tr>
					<td align="center" width="5" class="ClearPixel" height="2">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
					<td valign="top" width="100%">
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td valign="top">
                      <div class="frmbody" id="frm1" style="width: 600; height: 250">
                        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                          <tr width=600>
                            <td class="ExcelHeaderCell" align="center" width=10 >
                              <p align="center">S. No</td>
                            <td class="ExcelHeaderCell" align="center" >Item Name</td>
                            <td class="ExcelHeaderCell" align="center" >Purchase Type
                            </td>
						</tr>
						<%
						'Response.Write "<font color=#000000>"
						dim iCnt,ItemNode,ItemNm,sClassCode,sItemCode,i,sEntNo
						iCnt=0
						Set ItemNode=oNodRoot.Selectnodes("//ItemDetails/Item")
						for i=0 to ItemNode.Length-1
							iCnt=iCnt+1
							ItemNm = ItemNode.Item(i).Attributes.getNamedItem("ItmDescription").value
							sClassCode = ItemNode.Item(i).Attributes.getNamedItem("ClassificationCode").value
							sItemCode = ItemNode.Item(i).Attributes.getNamedItem("ItemCode").value
							sEntNo =  ItemNode.Item(i).Attributes.getNamedItem("EntryNo").value
						%>

						<tr>
							<td class="ExcelDisplayCell">
							<%=iCnt%>
							</td>
							<td class="ExcelDisplayCell">
							<%=ItemNm%>
							</td>
							<td class=exceldisplaycell>
							<select size="1" class="FormElem" name="cmbPurType<%=sClassCode%>Z<%=sItemCode%>Z<%=sEntNo%>">
								<option  value="" selected>Select</option>
								<%
									popSelPurTypeFull("0")
								%>
							</select>
							</td>
						</tr>
						<%Next%>
				       </table>
                      </div>
                    </td>
                  </tr>
                </table>
								</td>
								<td align="center" class="ClearPixel" width="6" height="2">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                                <tr>
								<td align="center" class="MiddlePack" colspan="3" width="100%">
								</td>
                                </tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<input type="button" value="Done" name="B8" onClick="Submit_Clk()"  class="ActionButton" tabindex="3" >
															<input type="button" value="Cancel" name="B10"  onClick="window.close()" class="ActionButton" tabindex="3" >
															<input type="reset" value="Reset" name="B9" class="ActionButton" tabindex="3" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="6">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                                <tr>
								<td align="center" class="BottomPack" colspan="3" width="100%">
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

