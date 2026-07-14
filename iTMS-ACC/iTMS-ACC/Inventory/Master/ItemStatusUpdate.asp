<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItemStatusUpdate.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran 
	'Created On					:	July 21,2011
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
    Dim objDOM,ndRoot,ndInv,ndChild
    Dim sQuery,sItemCode,sClassCode,sOrgCode,sStatus,sUserId
    
    sOrgCode = Session("organizationcode")
    sItemCode = Request("ItemCode")
    sClassCode = Request("ClassCode")
    sStatus = Request("Status")
    
    sUserid = getUserId
    
    con.begintrans
    
    if sStatus = "AC" then
        sQuery = "UPDATE INV_M_ITEMMASTER SET ITEMACTIVE = 'Y' WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        sQuery = "UPDATE INV_M_ITEMMASTER SET ITEMONHOLD = 0 WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        sQuery = "UPDATE INV_M_ITEMMASTER SET DeadStock='N' WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        
        sQuery = "INSERT INTO INV_T_ONHOLDDETAILS (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE,"&_
                 "ONHOLDREASON,HOLDRELEASEDON,HOLDRELEASEDBY) VALUES ("& sItemCode &","& Pack(sOrgCode) &","&_
                 ""& sClassCode &",NULL,CONVERT(DATETIME,GETDATE(),103),"& sUserId &")"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
    elseif sStatus = "IA" then
        sQuery = "UPDATE INV_M_ITEMMASTER SET ITEMACTIVE = 'N' WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        sQuery = "UPDATE INV_M_ITEMMASTER SET ITEMONHOLD = 0 WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        sQuery = "UPDATE INV_M_ITEMMASTER SET DeadStock='N' WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        sQuery = "INSERT INTO INV_T_ONHOLDDETAILS (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE,"&_
                 "ONHOLDREASON,HOLDRELEASEDON,HOLDRELEASEDBY) VALUES ("& sItemCode &","& Pack(sOrgCode) &","&_
                 ""& sClassCode &",NULL,CONVERT(DATETIME,GETDATE(),103),"& sUserId &")"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
    elseif sStatus = "OH" then
        sQuery = "UPDATE INV_M_ITEMMASTER SET ITEMONHOLD = 1 WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        sQuery = "UPDATE INV_M_ITEMMASTER SET DeadStock='N' WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
        sQuery = "INSERT INTO INV_T_ONHOLDDETAILS (ITEMCODE,ORGANISATIONCODE,CLASSIFICATIONCODE,"&_
                 "ONHOLDREASON,HOLDRELEASEDON,HOLDRELEASEDBY) VALUES ("& sItemCode &","& Pack(sOrgCode) &","&_
                 ""& sClassCode &",NULL,CONVERT(DATETIME,GETDATE(),103),"& sUserId &")"
        Response.Write "<p>"&sQuery
        con.execute sQuery
    elseif sStatus = "DS" then
        sQuery = "UPDATE INV_M_ITEMMASTER SET DeadStock='Y' WHERE ITEMCODE ="& sItemCode &" AND CLASSIFICATIONCODE = "& sClassCode &" AND ORGANISATIONCODE = '"& sOrgCode &"'"
        Response.Write "<p>"&sQuery
        con.execute sQuery
        
    end if' if sStatus = "AC" then
	
	
	if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End 
		Response.Clear 
		con.CommitTrans
	end if

	con.close
	set con = nothing
	Response.Redirect "ITEMLISTENTRY.ASP?ACTN=L"
 %>