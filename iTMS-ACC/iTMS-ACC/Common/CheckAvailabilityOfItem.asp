<%Option Explicit%>
<%
	'Program Name				:	CheckAvailabilityOfItem.asp
	'Module Name				:	
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:
%>
<!-- #include file="../include/DatabaseConnection.asp"-->
<!--#include file="../include/populate.asp"-->
<%
    Dim rsTemp
    Dim sCompanyItemCode,sItemDescription,sCallFrom,sValue,sQuery,sOrgCode
    set rsTemp =Server.CreateObject("ADODB.Recordset")
    sCallFrom = Request("CallFrom")
    sValue    = Request("Value")
    sOrgCode  = Request("OrgCode")
    
    if ucase(trim(sCallFrom)) = "ITEMCODE" then
        sQuery = "Select ItemCode from INV_M_ItemMaster where CompanyItemCode ="& pack(sValue) &" and OrganisationCode="& pack(sOrgCode)
    else
        sQuery = "Select ItemCode from INV_M_ItemMaster where ItemDescription = "& pack(sValue) &" and OrganisationCode="& pack(sOrgCode)
    end if
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        if UCase(Trim(sCallFrom))="ITEMCODE" then
            Response.Write "Item Code Already Exist"
        else
            Response.Write "Item Description Already Exist"
        end if
    else
        Response.Write "NO"
    end if
    rsTemp.Close 
%>

