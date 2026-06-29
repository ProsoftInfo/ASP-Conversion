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
<SCRIPT LANGUAGE=vbscript>
Function CheckSetup()
    if document.formname.chkCategory.checked = true then
        document.formname.selCategory.disabled = false
    else
        document.formname.selCategory.disabled = true
    end if
    
    if document.formname.chkSubCategory.checked = true then
        document.formname.selSubCategory.disabled = false
    else
        document.formname.selSubCategory.disabled = true
    end if
        
    if document.formname.chkClassification.checked = true then
        document.formname.selSubCategory.multiple= false
        document.formname.selSubCategory.selectedIndex = 0
        populateClassification 
        document.formname.selClassification.disabled = false
    else
        document.formname.selClassification.disabled = true
    end if
End Function
'*******************************
Function populateSubCategory()
    Dim objhttp,ndRoot,ndSubLevel
    Dim sCategoryCode,sSubCategory
    document.formname.chkSubCategory.checked = false
    document.formname.chkClassification.checked = false
    CheckSetup 
    if document.formname.selCategory.selectedIndex>-1 then
        sCategoryCode = document.formname.selCategory(document.formname.selCategory.selectedIndex).value
    end if 'if document.formname.selCategory.selectedIndex>-1 then
    'sSubCategory = document.formname.selSubCategory(document.formname.selSubCategory.selectedIndex).value
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","XmlGetClassification.asp?Level=0&Category="&sCategoryCode,false
    objhttp.send
    if Trim(objhttp.responseXML.xml)<>"" then
        SubCategory.loadXML(objhttp.responseXML.xml)
    end if
    set ndRoot = SubCategory.documentElement
    if ndRoot.haschildNodes() then
        document.formname.selSubCategory.length=0
        for each ndSubLevel in ndRoot.childNodes 
            if ndSubLevel.nodeName="Level" then
                document.formname.selSubCategory.length=document.formname.selSubCategory.length+1
                document.formname.selSubCategory(document.formname.selSubCategory.length-1).text = ndSubLevel.getAttribute("GroupName")
                document.formname.selSubCategory(document.formname.selSubCategory.length-1).value = ndSubLevel.getAttribute("GroupCode")
            end if
        next
    end if
End Function
'*******************************
Function populateClassification()
    Dim objhttp,ndRoot,ndSubLevel
    Dim sCategoryCode,sSubCategory
    'sCategoryCode = document.formname.selCategory(document.formname.selCategory.selectedIndex).value
    if document.formname.selSubCategory.selectedIndex >-1 then
        sSubCategory = document.formname.selSubCategory(document.formname.selSubCategory.selectedIndex).value
    end if
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","XmlGetClassification.asp?Level=1&SubCategory="&sSubCategory,false
    objhttp.send
    if Trim(objhttp.responseXML.xml)<>"" then
        Classification.loadXML(objhttp.responseXML.xml)
    end if
    set ndRoot = Classification.documentElement
    if ndRoot.haschildNodes() then
        document.formname.selClassification.length = 0
        for each ndSubLevel in ndRoot.childNodes 
            if ndSubLevel.nodeName="Level" then
                document.formname.selClassification.length = document.formname.selClassification.length + 1
                document.formname.selClassification(document.formname.selClassification.length-1).text = ndSubLevel.getAttribute("GroupName")
                document.formname.selClassification(document.formname.selClassification.length-1).value = ndSubLevel.getAttribute("GroupCode")
            end if
        next
    end if
End Function
'***************************************
Function AddCategorySetup()
Dim ndRoot,ndCategory,ndSubCategory,ndClassification,ndSelCategory,ndSelSubCategory,ndSelClassification
Dim sCategoryCode,sCategoryName,sCategoryConsider,sSubCateCode,sSubCateName,sSubCateConsider,sClassName,sClassCode,sClassConsider
Dim sSelCategory,sSelSubCategory,sSelClassification,sChkCategory,sChkSubCategory,sChkClassification
Dim sSelCategoryValue,sSelSubCategoryValue,sSelClassificationValue,sSelCategoryIndex,sSelSubCategoryIndex,sSelClassificationIndex
Dim bSelCategory,bSelSubCategory,bSelClassification

sChkCategory = document.formname.chkCategory.checked
sChkSubCategory = document.formname.chkSubCategory.checked 
sChkClassification = document.formname.chkClassification.checked 

if sChkCategory then 
    sSelCategoryIndex = document.formname.selCategory.selectedIndex 
    if sSelCategoryIndex > -1 then
        sSelCategoryValue  = document.formname.selCategory(sSelCategoryIndex).value
        sSelCategory = document.formname.selCategory(sSelCategoryIndex).text
    end if
end if

if sChkSubCategory then
    sSelSubCategoryIndex = document.formname.selSubCategory.selectedIndex 
    if sSelSubCategoryIndex > -1 then
        sSelSubCategoryValue = document.formname.selSubCategory(sSelSubCategoryIndex).value
        sSelSubCategory = document.formname.selSubCategory(sSelSubCategoryIndex).text
    end if
end if

set ndRoot = OutDataXML.documentElement
bSelCategory = false
bSelSubCategory = false
if ndRoot.hasChildNodes() then  
    for each ndCategory in ndRoot.childNodes
        if ndCategory.nodeName="Category" and trim(ndCategory.getAttribute("Code"))=trim(sSelCategoryValue) then 
            bSelCategory = true
            set ndSelCategory = ndCategory
            for each ndSubCategory in ndCategory.childNodes
                if ndSubCategory.nodeName="SubCategory" and Trim(ndSubCategory.getAttribute("Code"))=Trim(sSelSubCategoryValue) then    
                    bSelSubCategory = true
                    set ndSelSubCategory = ndSubCategory
                    exit for
                end if
            next
        end if
    next
end if

if not bSelCategory then
    set ndSelCategory = OutDataXML.createElement("Category") 
        ndSelCategory.setAttribute "Code",sSelCategoryValue 
        ndSelCategory.setAttribute "Name",sSelCategory 
        ndSelCategory.setAttribute "Consider","0"
        ndRoot.appendChild ndSelCategory 
end if
if not bSelSubCategory and sChkSubCategory then
    set ndSelSubCategory = OutDataXML.createElement("SubCategory")
    ndSelSubCategory.setAttribute "Code",sSelSubCategoryValue 
    ndSelSubCategory.setAttribute "Name",sSelSubCategory 
    ndSelSubCategory.setAttribute "Consider","0"
    ndSelCategory.appendChild ndSelSubCategory
end if

if sChkClassification then
    sSelClassificationIndex = document.formname.selClassification.selectedIndex 
    if sSelClassificationIndex>-1 then
        for iCnt = 0 to document.formname.selClassification.length - 1
           if document.formname.selClassification(iCnt).selected = true then
           bSelClassification = false
            sSelClassificationValue = document.formname.selClassification(iCnt).value
            sSelClassification = document.formname.selClassification(iCnt).text
                if ndSelSubCategory.hasChildNodes() then
                    for each ndClassification in ndSelSubCategory.childNodes
                        if ndClassification.nodeName="Classification" and Trim(ndClassification.getAttribute("Code"))=Trim(sSelClassificationValue) then
                            bSelClassification = true 
                        end if
                    next
                end if
                if not bSelClassification then
                    set ndSelClassification = OutDataXML.createElement("Classification")
                    ndSelClassification.setAttribute "Code",sSelClassificationValue 
                    ndSelClassification.setAttribute "Name",sSelClassification 
                    ndSelClassification.setAttribute "Consider","1"
                    ndSelSubCategory.appendChild ndSelClassification 
                end if
           end if
        next
    end if
else
    if sChkSubCategory then
        sSelSubCategoryIndex = document.formname.selSubCategory.selectedIndex 
        if sSelSubCategoryIndex>-1 then
            for iCnt = 0 to document.formname.selSubCategory.length - 1
                if document.formname.selSubCategory(iCnt).selected = true then
                    bSelSubCategory=false
                        sSelSubCategoryValue =document.formname.selSubCategory(iCnt).value
                        sSelSubCategory = document.formname.selSubCategory(iCnt).value
                        
                        if ndSelCategory.hasChildNodes() then
                            for each ndSubCategory in ndSelCategory.childNodes
                                if ndSubCategory.nodeName="SubCategory" and Trim(ndSubCategory.getAttribute("Code"))=Trim(sSelSubCategoryValue) then
                                    bSelSubCategory = true
                                    ndSubCategory.setattribute "Consider","1"
                                end if
                            next
                        end if
                        if not bSelSubCategory then
                            set ndSelSubCategory = OutDataXML.createElement("SubCategory")
                            ndSelSubCategory.setAttribute "Code",sSelSubCategoryValue 
                            ndSelSubCategory.setAttribute "Name",sSelSubCategory 
                            ndSelSubCategory.setattribute "Consider","1"
                            ndSelCategory.appendChild ndSelSubCategory
                        end if
                end if 'if document.formname.selSubCategory(iCnt).selected = true then
            next
        end if
    end if
end if
if not sChkSubCategory then
    ndSelCategory.setAttribute "Consider","1"
else
    if not sChkClassification then
        ndSelSubCategory.setAttribute "Consider","1"
    end if 
end if

DisplayTable
End Function
'******************************************
Function ClearTable()
	Dim i
	for	i = 1 to document.all.tblSetup.rows.length - 1
		document.all.tblSetup.deleteRow(1)
	next
End function
'*****************
Function DisplayTable()
Dim ndRoot,ndCategory,ndSubCategory,ndClassification
Dim sCategoryCode,sCategoryName,sCategoryConsider,sSubCateCode,sSubCateName,sSubCateConsider,sClassName,sClassCode,sClassConsider
Dim iSerNo
iSerNo = 0
    set ndRoot = OutDataXML.documentElement
    if ndRoot.hasChildNodes() then
        ClearTable()
        for each ndCategory in ndRoot.childNodes
            if ndCategory.nodeName="Category" then
                sCategoryCode = ndCategory.getAttribute("Code")
                sCategoryName = ndCategory.getAttribute("Name")
                sCategoryConsider = ndCategory.getAttribute("Consider")
                if sCategoryConsider =  "1" then
                    iSerNo = iSerNo + 1
                    
                    set tRow = document.all.tblSetup.insertRow(document.all.tblSetup.rows.length)
                    set sInnerCell = tRow.insertCell()
                    sInnerCell.className="ExcelSerial"
                    sInnerCell.innerHtml = iSerNo
                    sInnerCell.align = "Center"
                    
                    set sInnerCell = tRow.insertCell()
                    set oText = document.createElement("<input type=Checkbox name=ChkSetupZ"&iSerNo&" value='"&sCategoryCode&"::'>")
                    sInnerCell.appendChild(oText)
                    sInnerCell.className="ExcelDisplayCell"
                    sInnerCell.align = "Center"
                    
                    set sInnerCell = tRow.insertCell()
                    sInnerCell.className="ExcelDisplayCell"
                    sInnerCell.innerHtml = sCategoryName 
                    sInnerCell.align = "Left"
                    
                    set sInnerCell = tRow.insertCell()
                    sInnerCell.className="ExcelDisplayCell"
                    sInnerCell.innerHtml = "N/A"
                    sInnerCell.align = "Left"
                    
                    set sInnerCell = tRow.insertCell()
                    sInnerCell.className="ExcelDisplayCell"
                    sInnerCell.innerHtml = "N/A"
                    sInnerCell.align = "Left"
                else
                    if ndCategory.haschildNodes() then
                        for each ndSubCategory in ndCategory.childNodes
                            if ndSubCategory.nodeName="SubCategory" then
                                sSubCateCode = ndSubCategory.getAttribute("Code")
                                sSubCateName = ndSubCategory.getAttribute("Name")
                                sSubCateConsider= ndSubCategory.getAttribute("Consider")
                                if sSubCateConsider ="1" then
                                    iSerNo = iSerNo + 1
                    
                                    set tRow = document.all.tblSetup.insertRow(document.all.tblSetup.rows.length)
                                    set sInnerCell = tRow.insertCell()
                                    sInnerCell.className="ExcelSerial"
                                    sInnerCell.innerHtml = iSerNo
                                    sInnerCell.align = "Center"
                                    
                                    set sInnerCell = tRow.insertCell()
                                    set oText = document.createElement("<input type=Checkbox name=ChkSetupZ"&iSerNo&" value='"&sCategoryCode&":"& sSubCateCode &":'>")
                                    sInnerCell.appendChild(oText)
                                    sInnerCell.className="ExcelDisplayCell"
                                    sInnerCell.align = "Center"
                                    
                                    set sInnerCell = tRow.insertCell()
                                    sInnerCell.className="ExcelDisplayCell"
                                    sInnerCell.innerHtml = sCategoryName 
                                    sInnerCell.align = "Left"
                                    
                                    set sInnerCell = tRow.insertCell()
                                    sInnerCell.className="ExcelDisplayCell"
                                    sInnerCell.innerHtml = sSubCateName
                                    sInnerCell.align = "Left"
                                    
                                    set sInnerCell = tRow.insertCell()
                                    sInnerCell.className="ExcelDisplayCell"
                                    sInnerCell.innerHtml = "N/A"
                                    sInnerCell.align = "Left"
                                else
                                    if ndSubCategory.hasChildNodes() then
                                        for each ndClassification in ndSubCategory.childNodes
                                            if ndClassification.nodeName="Classification" then
                                                sClassCode = ndClassification.getAttribute("Code")
                                                sClassName = ndClassification.getAttribute("Name")
                                                sClassConsider = ndClassification.getAttribute("Consider") 
                                                if sClassConsider = "1" then
                                                    iSerNo = iSerNo + 1
                    
                                                    set tRow = document.all.tblSetup.insertRow(document.all.tblSetup.rows.length)
                                                    set sInnerCell = tRow.insertCell()
                                                    sInnerCell.className="ExcelSerial"
                                                    sInnerCell.innerHtml = iSerNo
                                                    sInnerCell.align = "Center"
                                                    
                                                    set sInnerCell = tRow.insertCell()
                                                    set oText = document.createElement("<input type=Checkbox name=ChkSetupZ"&iSerNo&" value='"&sCategoryCode&":"& sSubCateCode &":"& sClassCode &"'>")
                                                    sInnerCell.appendChild(oText)
                                                    sInnerCell.className="ExcelDisplayCell"
                                                    sInnerCell.align = "Center"
                                                    
                                                    set sInnerCell = tRow.insertCell()
                                                    sInnerCell.className="ExcelDisplayCell"
                                                    sInnerCell.innerHtml = sCategoryName 
                                                    sInnerCell.align = "Left"
                                                    
                                                    set sInnerCell = tRow.insertCell()
                                                    sInnerCell.className="ExcelDisplayCell"
                                                    sInnerCell.innerHtml = sSubCateName
                                                    sInnerCell.align = "Left"
                                                    
                                                    set sInnerCell = tRow.insertCell()
                                                    sInnerCell.className="ExcelDisplayCell"
                                                    sInnerCell.innerHtml = sClassName
                                                    sInnerCell.align = "Left"
                                                end if 'if sClassConsider = "1" then
                                            end if 'if ndClassification="Classification" then
                                        next
                                    end if
                                end if 'if sSubCateConsider ="1" then
                            end if 'if ndSubCategory.nodeName="SubCategory" then
                        next
                    end if 'if ndCategory.haschildNodes() then
                end if 'if sCategoryConsider =  "1" then
            end if
        next
        document.formname.hRow.value = iSerNo
    end if 'if ndRoot.hasChildNodes() then
End Function
'************************************************
Function LoadXML()
    Dim objhttp
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","XMLGetCategorySetup.asp",false
    objhttp.send
    if Trim(objhttp.responseXML.xml)<>"" then
        OutDataXML.loadXML(objhttp.responseXML.xml)
    end if
End Function
'*********************************
Function DeleteSetUp()
Dim nRow,sCategory,sSubCategory,sClassification,iSelCount,iCnt
Dim sArrCategory,ChkObj
Dim ndRoot,ndCategory,ndSubCategory,ndClassification
Dim sSelCategory,sSelClassification,sSelSubCategory
Dim iSelCateCons,iSelClassCons,iSelSubCateCons
    nRow = document.formname.hRow.value 
    iSelCount=0
    For iCnt = 1 to nRow
       set ChkObj = eval("document.formname.ChkSetupZ"&iCnt)
       if ChkObj.checked then 
            iSelCount = iSelCount + 1
            sArrCategory = Split(ChkObj.value,":")
       end if
    Next
    
    if iSelCount > 1 then
        alert("Select single entry to delete")
        exit function
    end if
    
    sCategory = sArrCategory(0)
    sSubCategory = sArrCategory(1)
    sClassification = sArrCategory(2)
    
    set ndRoot = OutDataXML.documentElement
    if ndRoot.haschildNodes() then
        for each ndCategory in ndRoot.childNodes
            if ndCategory.nodeName="Category" then
                sSelCategory = ndCategory.getAttribute("Code")
                iSelCateCons = ndCategory.getAttribute("Consider")
                if trim(sSelCategory) = trim(sCategory) then
                   if iSelCateCons = "1" then
                        ndCategory.setAttribute "Consider","0"
                   else
                        for each ndSubCategory in ndCategory.childNodes
                            if ndSubCategory.nodeName="SubCategory" then
                                sSelSubCategory = ndSubCategory.getAttribute("Code")
                                iSelSubCateCons = ndSubCategory.getattribute("Consider")
                                if Trim(sSelSubCategory)= Trim(sSubCategory) then
                                    if iSelSubCateCons="1" then
                                        ndSubCategory.setAttribute "Consider","0"
                                    else
                                        for each ndClassification in ndSubCategory.childNodes
                                            if ndClassification.nodeName="Classification" then
                                                sSelClassification = ndClassification.getAttribute("Code")
                                                iSelClassCons = ndClassification.getAttribute("Consider")
                                                if Trim(sSelClassification)=Trim(sClassification) then
                                                    if iSelClassCons = "1" then
                                                        ndClassification.setAttribute "Consider","0"
                                                    end if
                                                end if
                                            end if
                                        next
                                    end if
                                end if
                            end if
                        next
                   end if
                end if
            end if
        next
    end if
    DisplayTable
End Function
'***************************************
Function SetupCategory()
    Dim Objhttp
    set Objhttp = CreateObject("Microsoft.XMLHTTP")
    Objhttp.open "POST","ItemStockCloseSetupInsert.asp",false
    objhttp.send OutDataXML.XMLDocument
    if Trim(objhttp.responseText)<>"" then
        alert(objhttp.responseText)
    else
        alert("Records updated successfully")
        LoadXML 
        DisplayTable 
    end if
End Function
'**********************************************
</Script>
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