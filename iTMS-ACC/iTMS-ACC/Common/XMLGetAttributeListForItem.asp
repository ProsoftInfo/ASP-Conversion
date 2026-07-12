<%@language="VBScript"%>
<%Option Explicit%>
<%
	'Program Name				:	XMLGetAttributeListForItem.asp
	'Module Name				:	
	'Author Name				:	Ragavendran R
	'Created On					:	April 04,2012
	'Modified By				:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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

<!--#include file="../include/DatabaseConnection.asp"-->
<!--#include file="../include/populate.asp"-->
<!-- #include File="../include/CommonFunctions.asp" -->
<%
    Dim oDOM,objRs,rstemp
    Dim ndRoot,ndAttribute,ndOption
    Dim sQuery,sAttributeList
    Dim iItemCode
    set oDOM = Server.CreateObject("Microsoft.XMLDOM")
    set objRs = Server.CreateObject("ADODB.Recordset")
    set rstemp = Server.CreateObject("ADODB.Recordset")
    iItemCode = Request.QueryString("ItemCode")
    
    set ndRoot = oDOM.createElement("Root")
    oDOM.appendChild ndRoot
    
    sQuery = "Select isNull(AttributeList,0) from INV_M_ItemMaster where ItemCode= "& iItemCode &""
    objRs.Open sQuery,con
    if not objRs.EOF then
        sAttributeList =  objrs(0)
    end if
    objRs.Close 
    
    sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID in ("& sAttributeList &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName"
    objRs.Open sQuery,con
    if not objRs.EOF then
        do while not objRs.EOF 
            set ndAttribute = oDOM.createElement("Attribute")
                ndAttribute.setAttribute "ID",objRs(0)
                ndAttribute.setAttribute "Name",objRs(1)
                ndRoot.appendChild ndAttribute
                
                sQuery = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID = "& objRs(0)
                rstemp.Open sQuery,con
                if not rstemp.EOF then
                    do while not rstemp.EOF 
                        set ndOption = oDOM.createElement("Option")
                        ndOption.setAttribute "Value",rstemp(2)
                        ndOption.setAttribute "Name",rstemp(3)
                        ndAttribute.appendChild ndOption
                        rstemp.MoveNext  
                    loop
                end if
                rstemp.Close 
            objRs.MoveNext 
        loop
    end if
    objRs.Close 
    
    Response.ContentType = "text/xml"
    Response.Write oDOM.xml
%>