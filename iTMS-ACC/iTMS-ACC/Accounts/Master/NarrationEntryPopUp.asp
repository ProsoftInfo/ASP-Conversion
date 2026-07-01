<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	Narration.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	SENTHIL E
	'Created On					:	September 15
	'Modified by				:	UmaMaheswari S
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
<%	
	Dim oDOM,Root,NewElem,sNarrDesc,sNarrShortDesc,nSelBookCode,nSelBookNo,sBookName
	Dim nNarrNo,sBookCode,sType
		
	Dim objRs,sQuery,sFlag
	
	Set objRs = Server.CreateObject("ADODB.RecordSet")
	set oDOM = Server.CreateObject("Microsoft.XMLDom")
	
	sBookCode = Trim(Request.QueryString("BookCode"))
	sType = Trim(Request.QueryString("Type"))
	
	If sType = "E" Then
		nNarrNo = Trim(Request.QueryString("NarrNo"))
	End IF
	
	If sType = "E" Then
	
		sQuery="select NarrationShortDesc,NarrationDesc from Acc_M_FrequentDescriptions where NarrationNumber="&nNarrNo
		
		with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		
		set objRs.ActiveConnection=nothing
		
		if not objRs.EOF then
			sNarrShortDesc = objRs(0)
			sNarrDesc	   = objRs(1)
		end if
		objRs.Close
		
		with objRs
			.CursorLocation =3
			.CursorType =3
			.Source = "Select Distinct a.BookCode,a.BookNumber,b.BookName From Acc_R_BookFreqDesc a,Acc_R_ApplicableAccountHeads b Where a.NarrationNumber="& nNarrNo &" and a.BookCode=b.BookCode and a.BookNumber = b.BookNumber"
			.ActiveConnection = con
			.Open
		end with
		
		set objRs.ActiveConnection=nothing
		
		if not objRs.EOF then
			Do While Not objRs.EOF 
				
				nSelBookCode  = nSelBookCode & "," & objRs(0)
				nSelBookNo	  = nSelBookNo & "," & objRs(1)
				sBookName	  = sBookName & "," & objRs(2)
				objRs.MoveNext 
			Loop
		end if
		objRs.Close
		
		If nSelBookCode <> "" Then nSelBookCode = Mid(nSelBookCode,2)
		If nSelBookNo <> "" Then nSelBookNo = Mid(nSelBookNo,2)
		If sBookName <> "" Then sBookName = Mid(sBookName,2)
		
		set Root = oDOM.CreateElement("Root")
		
		set NewElem = oDOM.createElement("Desc")
		NewElem.setAttribute("Type"),Trim(sType)
		NewElem.setAttribute("ShortDesc"),Trim(sNarrShortDesc)
		NewElem.setAttribute("Desc"),Trim(sNarrDesc)
		NewElem.setAttribute("BookCode"),Trim(nSelBookCode)
		NewElem.setAttribute("BookNo"),Trim(nSelBookNo)
		NewElem.setAttribute("NarrNo"),Trim(nNarrNo)
		Root.appendChild NewElem
		
		oDOM.appendchild Root
		oDOM.save Server.MapPath("../Temp/Master/NarrationDet.xml")
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%If sType = "N" Then%>
	<XML ID="OutData"><Root/></XML>
<%Else%>
	<XML ID="OutData" src="<%="../Temp/Master/NarrationDet.xml"%>"></XML>
<%End If%>
<XML ID="RetData"><ROOT Done=""/></XML>
<XML ID="BookDet"><Root/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript src="../../scripts/Selection.js"></SCRIPT>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/ModalReturnCompat.js"></script>
<script language="javascript">
window.__itmsPopupCompat = { type: "narrationEntryPopup" };
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.xmlIsland("RetData");
});
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname" action="" >
<input type=hidden name="hSelectedValue" value="" >
<input type=hidden name="hActionFlag" value="N" >
<input type=hidden name="hBookCode" value="<%=sBookCode%>" >
<%If sType = "N" Then%>
	<input type=hidden name="hSelBookCode" value="" >
	<input type=hidden name="hSelBookNo" value="" >
	<input type=hidden name="hNarrNo" value="" >
<%Else%>
	<input type=hidden name="hSelBookCode" value="<%=nSelBookCode%>" >
	<input type=hidden name="hSelBookNo" value="<%=nSelBookNo%>" >
	<input type=hidden name="hNarrNo" value="<%=nNarrNo%>" >
<%End IF%>
<input type=hidden name="hType" value="<%=sType%>" >
    
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Narrations
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
											<td class="FieldCell" width="135">Short Description</td>
											<td>
												<input type="text" name="txtShortDesc" size="11" value="<%=sNarrShortDesc%>" maxlength="10" class="FormElem">
											</td>
										</tr>
                                
										<tr>
											<td class="FieldCell" width="135"> Description</td>
											<td>
												<input type="text" name="txtDesc" size="25" value="<%=sNarrDesc%>"  MaxLength=200 class="FormElem">
											</td>
										</tr>
										
										<tr>
											<td class="FieldCell" width="135"> Used In Books</td>
											<td>
												<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" alt="Select Used In Books" onclick="selBook()" width="15" height="15">
												<span class=DataOnly ID=UsedInBook><%=sBookName%></Span>
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
																<!--<input type="button" value="Delete" name="del" class="ActionButton" onClick="checkSubmit('D')">-->
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
