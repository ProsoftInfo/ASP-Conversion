<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItemInvInsert.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran 
	'Created On					:	July 18,2011
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
    Dim objDOM,ndRoot,ndInv,ndChild,rstemp
    Dim sQuery,sItemCode,sClassCode,sOrgCode,sABC,sVED,sFSN,sACC,nROQ,nROL,nEOQ
    Dim sFastMovCri,sSlowMovCri,sNonMovCri
    
    set objDOM = Server.CreateObject("Microsoft.XMLDOM")
    set rstemp = Server.CreateObject("ADODB.Recordset")
    
    objDOM.load(Request)
    objDOM.SAVE(Server.MapPath("../temp/Master/ItemInvStatus"&Session.SessionID&".xml"))
    
    con.begintrans
    
    set ndRoot = objDOM.documentElement
    sItemCode = ndRoot.getAttribute("ItemCode")
    sClassCode = ndRoot.getAttribute("ClassCode")
    sOrgCode = ndRoot.getAttribute("OrgCode")
    
    if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            if ndChild.nodeName="Inventory" then
                set ndInv = ndChild
            end if
        next
    end if
    
    sABC = ndInv.getAttribute("ABC")
    sVED = ndInv.getAttribute("VED")
    sACC = ndInv.getAttribute("ACC")
    sFSN = ndInv.getAttribute("FSN")
    sFastMovCri = ndInv.getAttribute("Fast")
    sSlowMovCri = ndInv.getAttribute("Slow")
    sNonMovCri = ndInv.getAttribute("Non")
    nROL = ndInv.getAttribute("RL")
    nROQ = ndInv.getAttribute("RQ")
    nEOQ = ndInv.getAttribute("EQ")
    if Trim(sABC)="" or IsNull(sABC) then sABC = "NULL"
    if trim(sVED)="" or IsNull(sVED) then sVED = "NULL"
    if trim(sFSN)="" or IsNull(sFSN) then sFSN = "NULL"
    
    if sABC<>"NULL" then sABC = pack(sABC)
    if sVED<>"NULL" then sVED = pack(sVED)
    if sFSN<>"NULL" then sFSN = pack(sFSN)
    
    if Trim(sFastMovCri)="" or IsNull(sFastMovCri) then sFastMovCri = "0"
    if Trim(sSlowMovCri)="" or IsNull(sSlowMovCri) then sSlowMovCri = "0"
    if Trim(sNonMovCri)=""  or IsNull(sNonMovCri) then sNonMovCri = "0"
    
    sQuery = "Select ItemCode from INV_M_ITEMORGINVENTORY where ItemCode ="& sItemCode &" and ClassificationCode = "& sClassCode &" and OrganisationCode ="& sOrgCode
    rstemp.Open sQuery,con
    if rstemp.EOF then
        sQuery = "Insert into INV_M_ITEMORGINVENTORY (ItemCode,ClassificationCode,OrganisationCode,ABCClassification,FSNCategory,VEDCategory,FastMovingCriteria,SlowMovingCriteria,NonMovingCriteria)"&_
                 " values("&sItemCode&","& sClassCode&","& Pack(sOrgCode) &","& sABC &","& sFSN &","& sVED&","& sFastMovCri &","& sSlowMovCri & ","& sNonMovCri &")"
    else
        sQuery = "Update INV_M_ITEMORGINVENTORY set ABCClassification="& sABC &",FSNCategory="& sFSN &",VEDCategory="& sVED&",FastMovingCriteria="& sFastMovCri &",SlowMovingCriteria="& sSlowMovCri &",NonMovingCriteria ="& sNonMovCri  &""&_
            " where ItemCode ="&sItemCode&" and ClassificationCode = "& sClassCode&" and OrganisationCode ="& Pack(sOrgCode) 
      
    end if
    rstemp.Close 
    
    Response.Write "<p>"& sQuery
    con.execute sQuery
    
    sQuery = "Update INV_M_ItemMaster set ReOrderLevel="& nROL &",ReOrderQty="& nROQ &",EcoOrderQty="& nEOQ &" where ItemCode = "& sItemCode &" and ClassificationCode = "& sClassCode &" and OrganisationCode = "& pack(sOrgCode)
    Response.Write "<p>"& sQuery
    con.execute sQuery
    
    
	'		sSql = "INSERT INTO INV_M_ITEMORGINVENTORY (ITEMCODE,CLASSIFICATIONCODE,ORGANISATIONCODE,VALUATIONMETHOD,SHELFLIFEOFITEM," &_
	'			"AVGCONSUMPFREQ,ALLOWLOCATIONTRANSFER,ALLOWINTERUNITTRANSFER,ALLOWREPLENISHMENT, " &_
	'			"ABCCLASSIFICATION,FSNCATEGORY,VEDCATEGORY,STOCKHOLDINGPERIOD," &_
	''			"ALLOWCYCLECOUNT,CYCLECOUNTFREQUENCY,REORDERLEVEL,REORDERQUANTITY, " &_
	'			"ECONOMICORDERQTY,ITEMDEFINEDBY,ITEMDEFINEDON) VALUES " &_
	'			"(" & iItmCode & "," & sClassCode & "," & MyPack(sOrgCode) & "," & MyPack(sAccType) & ",0," &_
	'			"0," & MyPack(sLocEli) & "," & MyPack(sIntEli) & "," & MyPack(sRepEli) & "," &_
	'			"" & MyPack(sABCValue) & "," & MyPack(sFSNValue) & "," & MyPack(sVED) & "," & sStock & "," &_
	'			"'0',1," & iReLvl & "," & iReQty & "," &_
	'			"" & iEcQty & "," & iController & ",CONVERT(DATETIME,GETDATE(),103))"
	
	
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
 %>