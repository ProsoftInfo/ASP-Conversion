<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmSpecsEditPop.asp
	'Module Name				:	Inventory (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jul 29,2011
	'Modified By                :   
	'Modified On				:   
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
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
    Dim sQuery,rsTemp,sItemTypeAttID,sItemTypeAttName,sDataLength,iCnt
    Dim sItemType,iCtr,sClassCode,sValue,sItemCode
    Dim oDOM,ndRoot,ndAttType,ndOption,ndHeader
    Dim dcrs1,dcrs2,dcrs,iOptVal,sOptName
    
    Set rsTemp = Server.CreateObject("ADODB.RecordSet")
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")
    set dcrs2 = Server.CreateObject("ADODB.Recordset")
    set oDOM = Server.CreateObject("Microsoft.XMLDOM")
    
    Response.Write "<font color=red>"
    'Response.Write Request.QueryString
    
    sItemType = Request("ItemType")
    sClassCode = Request("ClassCode")
    sItemCode = Request("ItemCode")
    if trim(sItemType)="" then sItemType = "STO"
    
    set ndRoot = oDOM.createElement("ItemSpecs")
    oDOM.appendChild ndRoot
    
    with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT HEADERID,ITEMTYPEHEADERNAME FROM INV_M_ITEMTYPEHEADER WHERE HEADERID = 6 ORDER BY HEADERID"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing
	if not dcrs1.EOF then
	    iCtr = 0
	    set ndHeader = oDOM.createElement("TypeHeader")
	        ndHeader.setAttribute "ID",dcrs1(0)
	        ndHeader.setAttribute "Name",dcrs1(1)
        ndRoot.appendChild ndHeader
			with dcrs2
				.CursorLocation = 3
				.CursorType = 3
				'.Source = "SELECT ITEMTYPEATTRIBUTEID,HEADERID,ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE,ITEMTYPEATTRIBUTEDATALENGTH,ITEMTYPEATTRIBUTEDECIMAL FROM INV_M_ITEMTYPEATTRIBUTES WHERE HEADERID = 6 AND ITEMTYPEID = " & Pack(sItemType) & " and ClassificationCode = "& sClassCode &" ORDER BY ITEMTYPEATTRIBUTEID"
				.Source = "SELECT ITEMTYPEATTRIBUTEID,HEADERID,ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE,ITEMTYPEATTRIBUTEDATALENGTH,ITEMTYPEATTRIBUTEDECIMAL FROM INV_M_ITEMTYPEATTRIBUTES WHERE HEADERID = 6 AND ClassificationCode = "& sClassCode &" ORDER BY ITEMTYPEATTRIBUTEID"
				.ActiveConnection = con
				.Open
			end with
			set dcrs2.ActiveConnection = nothing
			if not dcrs2.EOF then
				do while not dcrs2.EOF
					iCtr = iCtr + 1
					
					sQuery = "Select AttributeValue from INV_M_ItemMasterAttributes where ItemCode = "& sItemCode &" and "&_
         					 " ClassificationCode = "& sClassCode &" and HeaderID = 6 and ItemTypeAttributeID = "& dcrs2(0)
					rsTemp.Open sQuery,con
					if not rsTemp.EOF then
					    sValue = rsTemp(0)
					end if
					rsTemp.Close 
					
				    set ndAttType = oDOM.createElement("ATTRIBUTE")
				        ndAttType.setAttribute "NO",iCtr
				        ndAttType.setAttribute "ID",lcase(trim(dcrs2(0)))
				        ndAttType.setAttribute "NAME",lcase(trim(dcrs2(2)))
				        ndAttType.setAttribute "TYPE",lcase(trim(dcrs2(3)))
				        ndAttType.setAttribute "VALUE",sValue 
				    ndHeader.appendChild ndAttType
				        
	                    if lcase(trim(dcrs2(3))) = "options" then 
	                            with dcrs
			                        .CursorLocation = 3
			                        .CursorType = 3
			                        .Source = "SELECT OPTIONVALUE,OPTIONNAME FROM INV_M_ITEMTYPEOPTIONS WHERE ITEMTYPEATTRIBUTEID = " & trim(dcrs2(0)) & " ORDER BY OPTIONVALUE"
			                        .ActiveConnection = con
			                        .Open
		                        end with
		                        set dcrs.ActiveConnection = nothing
		                        set iOptVal = dcrs(0)
		                        set sOptName = dcrs(1)

		                        Do While Not dcrs.EOF
		                            set ndOption = oDOM.createElement("Option")
		                                ndOption.setAttribute "Value",iOptVal
		                                ndOption.setAttribute "Name",sOptName 
		                                ndOption.setAttribute "Selected","N"
		                            ndAttType.appendChild ndOption
			                        dcrs.MoveNext
		                        Loop
		                        dcrs.Close
					    end if
					dcrs2.MoveNext
				loop
			end if
			dcrs2.Close
	end if
	dcrs1.Close
	oDOM.save(Server.MapPath("../temp/master/TempItemSpecs_"&Session.SessionID&".xml"))
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Specifications</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" id="ItemSpecData" data-itms-xml-island="1" data-src="<%="../temp/master/TempItemSpecs_"&Session.SessionID&".xml"%>"><ItemSpecs/></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
window.ITMS_ITEM_SPECS_EDIT = true;
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemSpecsPopup.js"></SCRIPT>

</HEAD>
 
<BODY leftMargin=0 topMargin=0 onload="Init()">
<form method="POST" name="formname" action="">
<input type=hidden name=hCnt value="">
<input type=hidden name="hItemCode" value="<%=sItemCode%>">
<input type=hidden name="hClassCode" value="<%=sClassCode%>">
<input type=hidden name="hItemType" value="<%=sItemType%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Item Specs.</p>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="1" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
								    <td align="center" class="MiddlePack">
								    </td>
                                </tr>
							    <tr>
								    <td valign="top">
								        <table id="tblItemData" border="0" cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
									    </table>
								    </td>
							    </tr>
							    <tr>
								    <td align="center" class="MiddlePack">
								    </td>
                                </tr>
                                <tr>
								    <td align="center" class="ActionCell">
								        <input type=button name="btnSave" class="ActionButtonX" value="Save" onclick="CheckSubmit()">
								        <input type=button name="btnClose" class="ActionButtonX" value="Close" onclick="window.close()">
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
