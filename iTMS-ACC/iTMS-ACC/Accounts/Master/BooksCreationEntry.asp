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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
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

<SCRIPT LANGUAGE=javascript>
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
<XML ID="SeriesNoData" src="../../NoSeries/xmldata/SeriesNumberDetail.xml"></XML>
<xml id="UnitBook"><Root /></xml>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<SCRIPT language="vbscript">
'''''''''''''''''''''''''''''''''''''''
Function popUnitBooks()
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.Open "GET","XMLGetDayBooks.asp",false
    objhttp.send
    if trim(objhttp.responseXML.xml)<>"" then
        UnitBook.loadXML(objhttp.responseXML.xml)
    else
        alert(objhttp.responseText)
    end if

    objhttp.open "POST","XMLSaveParty.asp?Name=Unit&Mod=Book",false
    objhttp.send UnitBook.xml
    if trim(objhttp.responseText)<>"" then
        alert(objhttp.responseText)
    end if
End Function
'''''''''''''''''''''''''''''''''''''''''
Function setPayRec()
	if document.formname.selDayBook.value="01" or document.formname.selDayBook.value="02" then
		document.formname.selPayRecNo.options.length = 0
		document.formname.selPayRecNo.length = 2
		document.formname.selPayRecNo.options(0).text = "Yes"
		document.formname.selPayRecNo.options(0).Value = "Y"


		document.formname.selPayRecNo.options(1).text = "No"
		document.formname.selPayRecNo.options(1).Value = "N"
	else
		document.formname.selPayRecNo.options.length = 0
		document.formname.selPayRecNo.length = 1
		document.formname.selPayRecNo.options(0).text = "No"
		document.formname.selPayRecNo.options(0).Value = "N"
	end if
end Function

Function DisplayBook()
	dim iSeriesNo
	dim Root,sExp,iEntLen
	Set Root = SeriesNoData.documentElement
	ClearTable document.formname.selPayRecNo.value
	j=1
	if document.formname.selNoSeries.selectedIndex<> "0" then
		iSeriesNo= document.formname.selNoSeries.value
		sExp = "//Series[@No="&iSeriesNo&"]/Entry"
		Set TempNode = Root.selectNodes(sExp)
		iEntLen = Tempnode.length

		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iSeriesNo then
				IF CInt(iEntLen) = 12 Then
					document.formname.hSeriesType.value="M"
				Elseif Cint(iEntLen) = 4 Then
					document.formname.hSeriesType.value="Q"
				Elseif CInt(iEntLen) = 1 Then
					document.formname.hSeriesType.value="Y"
				Else
					document.formname.hSeriesType.value=HeaderNode.Attributes.Item(2).nodeValue
				End IF
				document.formname.hSeriesLen.value=HeaderNode.Attributes.Item(4).nodeValue

				For Each EntryNode In HeaderNode.childNodes

					iEntryNo=EntryNode.Attributes.Item(0).nodeValue
					sPeriod=EntryNode.Attributes.Item(1).nodeValue
					iNumber=EntryNode.Attributes.Item(2).nodeValue
					sPrefix=EntryNode.Attributes.Item(3).nodeValue
					sSufix=EntryNode.Attributes.Item(4).nodeValue

					Select Case HeaderNode.Attributes.Item(3).nodeValue
					   Case "M" sPeriod="Month-"&sPeriod
					   Case "Q" sPeriod="Quater-"&sPeriod
					   Case "Y" sPeriod="Yearly"
					End Select
					if document.formname.selPayRecNo.value="Y" then
						set oRow = document.all.tblBook.insertRow(j)

						InsertCell oRow,1,"",j,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sPeriod,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,2,"txtCrStartNo"&iEntryNo,iNumber,"ExcelInputCell","","",5,4,0,0,""
						InsertCell oRow,2,"txtCrPrefix"&iEntryNo,sPrefix,"ExcelInputCell","","",12,11,0,0,""
						InsertCell oRow,2,"txtCrSuffix"&iEntryNo,sSufix,"ExcelInputCell","","",11,10,0,0,""
						InsertCell oRow,2,"txtDrStartNo"&iEntryNo,iNumber,"ExcelInputCell","","",5,4,0,0,""
						InsertCell oRow,2,"txtDrPrefix"&iEntryNo,sPrefix,"ExcelInputCell","","",12,11,0,0,""
						InsertCell oRow,2,"txtDrSuffix"&iEntryNo,sSufix,"ExcelInputCell","","",11,10,0,0,""
					else

						set oRow = document.all.tblBook.insertRow(j)
						InsertCell oRow,1,"",j,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sPeriod,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,2,"txtStartNo"&iEntryNo,iNumber,"ExcelInputCell","","",5,4,0,0,""
						InsertCell oRow,2,"txtPrefix"&iEntryNo,sPrefix,"ExcelInputCell","","",12,11,0,0,""
						InsertCell oRow,2,"txtSuffix"&iEntryNo,sSufix,"ExcelInputCell","","",11,10,0,0,""
					end if
					j=j+1
				next
			end if
		next
	end if
end Function

Function popDayBookList()
dim iUnitNo,sUnitName

iUnitNo= document.formname.hUnitId.value
sUnitName=document.formname.hUnitName.value

'if document.formname.selUnitId.selectedIndex >0 then
	showModalDialog "PopBookNarrationList.asp?orgid="+iUnitNo+"&Mod=H","","dialogHeight:450px;dialogWidth:380px;center:Yes;help:No;resizable:No;status:No"

'else
'	MsgBox ("Select Unit")
'	document.formname.selUnitId.focus
'end if
end Function

Function popSeriesNo()
	Set Root = SeriesNoData.documentElement

	document.formname.selNoSeries.options.length = 0
	document.formname.selNoSeries.length = document.formname.selNoSeries.length+1
	document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).text = "Select Number Series"
	document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).Value = "0"

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(3).nodeValue = "M" then
			document.formname.selNoSeries.length = document.formname.selNoSeries.length+1
			document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
			document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).Value =HeaderNode.Attributes.Item(0).nodeValue
		end if
	next
end Function

Function ClearTable(sFlag)
	dim i
	for i=0 to document.all.tblBook.rows.length - 1
		document.all.tblBook.deleteRow(0)
	next
	if sFlag="Y" then
		set oRow = document.all.tblBook.insertRow(0)
		InsertCell oRow,1,"","S.No","ExcelSerial","Center","",0,0,0,0,""
		InsertCell oRow,1,"","Period","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","CR StartNo","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","CR Prefix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","CR Suffix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","DR StartNo","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","DR Prefix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","DR Suffix","ExcelHeaderCell","left","",0,0,0,0,""
	else
		set oRow = document.all.tblBook.insertRow(0)
		InsertCell oRow,1,"","S.No","ExcelSerial","Center","",0,0,0,0,""
		InsertCell oRow,1,"","Period","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","StartNo","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","Prefix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","Suffix","ExcelHeaderCell","left","",0,0,0,0,""
	end if
end Function

Function validateForm()
'	if document.formname.selUnitId.selectedIndex=0 then
'		MsgBox "Select Unit"
'		 document.formname.selUnitId.focus
'		Exit Function
'	end if
	if document.formname.selDayBook.selectedIndex=0 then
		MsgBox "Select Day Book"
		 document.formname.selDayBook.focus
		Exit Function
	end if
	if Trim(document.formname.txtName.value)="" then
		MsgBox "Enter Book Name"
		 document.formname.txtName.focus
		Exit Function
	end if
	if document.formname.selNoSeries.selectedIndex=0 then
		MsgBox "Select No Series"
		 document.formname.selNoSeries.focus
		Exit Function
	end if


	document.formname.B2.disabled = True
	document.formname.B4.disabled = True
	document.formname.submit
end Function
'''''''''''''''''''''''''''
Function CheckNumberSerious()
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","../../Admin/Master/XMLGetNoSeriesPattern.asp",false
    objhttp.send
    if trim(objhttp.responseXML.xml)<>"" then
        SeriesNoData.LoadXML(objhttp.responseXML.xml)
    else
        alert(objhttp.responseText)
    end if
End Function
</script>
<script language="javascript">
window.__itmsPopupCompat = { type: "booksCreationEntry" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
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
				       	             <a href="javascript:popDayBookList()"><span style="cursor: hand" Title="View Contra Details" >
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
