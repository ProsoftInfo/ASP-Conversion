<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	SetupInvBooks.asp
	'Module Name				:	Inventory
	'Author Name				:	Ragavendran R
	'Created On					:	May 06,2014
	'Modified By				:	
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/GetSerialDetail.asp" -->
<%
Dim sOrgID,objFSO
sOrgID = Session("organizationcode")
set objFSO = Server.CreateObject("Scripting.FileSystemObject")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Serial Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="UnitBookData"><Root></Root></script>
<%
if objFSO.FileExists(server.MapPath("../XMLData/BookSetup.xml")) then
%>
<script type="application/xml" data-itms-xml-island="1" id="BookXML" data-src="<%="../XMLData/BookSetup.xml"%>"></script>
<%
else
%>
<script type="application/xml" data-itms-xml-island="1" id="BookXML"><Root></Root>
</script>
<%
end if
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/selection.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function Init()
    sOrgCode= document.formname.hOrgID.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","../../Accounts/Transaction/XMLGetOrgBook.asp?BkCode=08&orgID=" & sOrgCode , false
	objhttp.send
	if objhttp.responseXML.xml <> "" then
		UnitBookData.loadXML objhttp.responseXML.xml
		Set Root = UnitBookData.documentElement
		For Each HeaderNode In Root.childNodes
			document.formname.selBooks.length = document.formname.selBooks.length+1
			document.formname.selBooks.options(document.formname.selBooks.length-1).text = HeaderNode.getAttribute("BookName")
			document.formname.selBooks.options(document.formname.selBooks.length-1).Value = HeaderNode.getAttribute("BookNumber")
		next
	end if
	set ndRoot = BookXML.documentElement
	if ndRoot.hasChildNodes() then
	    for each ndChild in ndRoot.childNodes
	        nBookNo = ndChild.getAttribute("No")
	        sBookName= ndChild.getAttribute("Name")
	        exit for 
	    next
	end if 
	
	set objBook = eval("document.formname.selBooks")
	if Trim(nBookNo)="" or IsNull(nBookNo) then nBookNo = ""
	if Trim(nBookNo)<>"" then
	    for iCnt = 0 to objBook.length-1
	        if trim(nBookNo)=trim(objBook(iCnt).value) then
	            objBook.selectedIndex = iCnt
	            exit for
	        end if
	    next
	end if 
End Function
'**************************************
Function CheckSubmit()
Dim objBook,nBookNo,nBookName
set objBook = eval("document.formname.selBooks")
    if objBook.selectedIndex<0 then
        alert("Select the Book")
        objBook.focus
        exit function
    end if 
    nBookNo = objBook(objBook.selectedIndex).value
    sBookName = objBook(objBook.selectedIndex).text
    set ndRoot = BookXML.documentElement
    if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            if trim(ndChild.nodeName)="Book" then
                ndChild.setAttribute "No",nBookNo
                ndChild.setAttribute "Name",sBookName
            end if 
        next
    else
        set ndChild = BookXML.createElement("Book")
        ndChild.setAttribute "No",nBookNo
        ndChild.setAttribute "Name",sBookName
        ndRoot.appendChild ndChild
    end if 

    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "POST","SetupInvBookInsert.asp",false
    objhttp.send BookXML.XMLDocument
    if objhttp.responseText<>"" then
        alert(objhttp.responseText)
    else
        alert("Book Setup is done")
    end if
End Function
</SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hOrgID" value="<%=sOrgID%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Book Selection
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
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%" align="center">
                                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                       <tr>
                                        <td class="FieldCellSub" align="center">
                                            <select id="selBooks" size="10" class="FormElem">
                                                
                                            </select>
                                        </td>
                                       </tr>
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Close" name="btnClose" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack"></td>
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
