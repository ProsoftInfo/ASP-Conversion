<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ManageVouchers_MoveBookPop.asp
	'Module Name				:	ACCOUNTS - Manage Vouchers
	'Author Name				:	KalaiSelvi R
	'Created On					:	September 24,2011
	'Modified by				:	
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%	
	Dim objRs,sQuery,sCallFrom,sTransNo,sBookCode
	Dim sExistingBookNo,sExistingBookName,sExistingBookAccHead 
	
	'Response.Write "<p> Request.QueryString = " & Request.QueryString
	
	sCallFrom = Request.QueryString("CallFrom")
	sTransNo = Request.QueryString("TransNo")
	
	
	if trim(sCallFrom) = "CASH"  then
		sBookCode = "01"
	else
		sBookCode = "02"
	end if 		
	
	Set objRs = Server.CreateObject("ADODB.RecordSet")	
		
	sQuery="select v.BookNumber,Upper(v.BookName),isnull(v.BookAccountHead,0),v.OtherUnitTransaction"&_
		" from vwOrgBookNames v,ACC_T_CREATEDVOUCHERHEADER a " &_
		" where a.BookCode=v.BookCode and a.BookNumber = v.BookNumber " &_
		" and a.CreatedTransNo = " & sTransNo & " and a.BOOKCODE=" & sBookCode & ""
		 
	if sBookCode="01" or sBookCode="02"  then
		sQuery=sQuery&" and BookAccountHead is not null "
	end if
	
	sQuery=sQuery&" Order By BookName "
		
	with objRs
		.CursorLocation =3
		.CursorType =3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
		
	set objRs.ActiveConnection=nothing
		
	if not objRs.EOF then
		sExistingBookNo		= objRs(0)
		sExistingBookName	= objRs(1)
		sExistingBookAccHead= objRs(2)
	end if 
	objRs.Close 

	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home - Change Book</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<script type="application/xml" data-itms-xml-island="1" ID="RetData"><ROOT Done="N"/></script>

<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT src="../../scripts/Selection.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script>
function checkSubmit()
{
	var frm = document.formname;
	var root = window.ITMSModalReturnCompat.xmlIsland("RetData");
	var parts;
	var newElem;

	if (frm.SelBook.selectedIndex === -1) {
		alert("Select Book");
		frm.SelBook.focus();
		return false;
	}

	parts = String(frm.SelBook.value || "").split("Z");
	newElem = root.ownerDocument.createElement("Book");
	newElem.setAttribute("TransNo", frm.hTransNo.value);
	newElem.setAttribute("ExistingBookNo", frm.hExistingBookNo.value);
	newElem.setAttribute("ExistingBookAccHead", frm.hExistingBookAccHead.value);
	newElem.setAttribute("NewBookNo", parts[0] || "");
	newElem.setAttribute("NewBookAccHead", parts[1] || "");
	root.appendChild(newElem);
	root.setAttribute("Done", "Y");
	window.ITMSModalReturnCompat.returnAndClose(root);
	return false;
}

window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.xmlIsland("RetData");
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="" >

<input type="hidden" name="hTransNo" value="<%=sTransNo%>">
<input type="hidden" name="hExistingBookNo"			value="<%=sExistingBookNo%>">
<input type="hidden" name="hExistingBookAccHead"	value="<%=sExistingBookAccHead%>">
    
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Move to Book
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
			<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
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
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                
                                    </table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                        <tr>
											<td class="FieldCell">Existing Book</td>
											<td>
												<Span class="DataOnly" id="SpnExistingBook"><%=sExistingBookName%></Span>
											</td>
										</tr>
                                		
										<tr>
											<td class="FieldCell">Select Book</td>
											<td>
												<Select name="SelBook" size=5 class="FormElem">
												<%
												
												sQuery="select v.BookNumber,Upper(v.BookName),isnull(v.BookAccountHead,0),v.OtherUnitTransaction"&_
													" from vwOrgBookNames v Where v.BookCode=" & sBookCode & " and v.BookNumber not in ( " & sExistingBookNo  & ")"	&_
													" and v.BookAccountHead not in (" & sExistingBookAccHead & ")"
														 
												if sBookCode="01" or sBookCode="02"  then
													sQuery=sQuery&" and BookAccountHead is not null "
												end if
	
												sQuery=sQuery&" Order By BookName "	
	
																										
 														
												with objRs
													.CursorLocation =3
													.CursorType =3
													.Source = sQuery
													.ActiveConnection = con
													.Open
												end with
													
												set objRs.ActiveConnection=nothing
													
												if not objRs.EOF then
													Do While Not objRs.EOF 
														%>
														<option value="<%=objRs(0)%>Z<%=objRs(2)%>"><%=objRs(1)%></option>
														<%															
														objRs.MoveNext 
													Loop
												end if
												objRs.Close
												%>
												</Select>
											</td>
										</tr>
                            
                                    </table>
								</td>
								<td align="center">
								</td>
							</tr>
                            								
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
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
                                                                <input type="button" value="Save" name="next" class="ActionButton" onClick="checkSubmit()">
                                                                <input type="button" value="Close" name="Close" class="ActionButton" onClick="window.close()">
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
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
