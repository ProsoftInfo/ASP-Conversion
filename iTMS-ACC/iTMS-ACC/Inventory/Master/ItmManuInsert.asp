<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmManuInsert.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran 
	'Created On					:	Feb 21,2013 
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
    Dim objDOM,ndRoot,ndInv,ndChild,rstemp,ndManu
    Dim sQuery,sItemCode,sClassCode,sOrgCode,sABC,sVED,sFSN,sACC
    Dim sFastMovCri,sSlowMovCri,sNonMovCri
    
    set objDOM = Server.CreateObject("Microsoft.XMLDOM")
    set rstemp = Server.CreateObject("ADODB.Recordset")
    
    objDOM.load(Request)
    objDOM.SAVE(Server.MapPath("../temp/Master/ItemManu"&Session.SessionID&".xml"))
    
    con.begintrans
    
    set ndRoot = objDOM.documentElement
    sItemCode = ndRoot.getAttribute("ItemCode")
    sClassCode = ndRoot.getAttribute("ClassCode")
    sOrgCode = ndRoot.getAttribute("OrgCode")
    
    if ndRoot.hasChildNodes() then
        for each ndChild in ndRoot.childNodes
            if ndChild.nodeName="Manufacture" then
                set ndManu = ndChild
            end if
        next
    end if
    
    Dim partycode,alias,cdn,grade,modelprocess,noofcavities,plateno,itemweight,mdn
    Dim sbasevalue,sitemrate,sexportrate,scurrency,patternmaterial,patternowner,patternavailability
    
    partycode   = ndManu.getAttribute("partycode")
    alias       = ndManu.getAttribute("alias")
    cdn         = ndManu.getAttribute("cdn")
    mdn         = ndManu.getAttribute("mdn")
    grade       = ndManu.getAttribute("grade")
    modelprocess= ndManu.getAttribute("modelprocess")
    noofcavities= ndManu.getAttribute("noofcavities")
    plateno     = ndManu.getAttribute("plateno")
    itemweight  = ndManu.getAttribute("itemweight")
    sbasevalue   = ndManu.getAttribute("basevalue")
    sitemrate    = ndManu.getAttribute("itemrate")
    sexportrate  = ndManu.getAttribute("exportrate")
    scurrency    = ndManu.getAttribute("currency")
    patternmaterial     = ndManu.getAttribute("patternmaterial")
    patternowner        = ndManu.getAttribute("patternowner")
    patternavailability = ndManu.getAttribute("patternavailability")
    
    if trim(partycode)="" or isNull(partycode) then partycode="0"
    
    if trim(alias)="" or isnull(alias) then alias ="NULL"
    if trim(alias)<>"NULL" then alias = pack(alias)
    
    if trim(cdn)="" or isnull(cdn) then cdn ="NULL"
    if trim(cdn)<>"NULL" then cdn = pack(cdn)
    
    if trim(mdn)="" or isnull(mdn) then mdn ="NULL"
    if trim(mdn)<>"NULL" then mdn = pack(mdn)
    
    if trim(grade)="" or isnull(grade) then grade ="NULL"
    if trim(grade)<>"NULL" then grade = pack(grade)
    
    if trim(modelprocess)="" or isnull(modelprocess) then modelprocess ="NULL"
    if trim(modelprocess)<>"NULL" then modelprocess = pack(modelprocess)
    
    if trim(noofcavities)="" or isnull(noofcavities) then noofcavities ="NULL"
    
    if trim(plateno)="" or isnull(plateno) then plateno ="NULL"
    
    
    if trim(itemweight)="" or isnull(itemweight) then itemweight ="NULL"
    
    if trim(sbasevalue)="" or isnull(sbasevalue) then sbasevalue ="NULL"
    
    if trim(sitemrate)="" or isnull(sitemrate) then sitemrate ="NULL"
    
    if trim(sexportrate)="" or isnull(sexportrate) then sexportrate ="NULL"
    
    if trim(scurrency)="" or isnull(scurrency) then scurrency ="NULL"
    
    if trim(patternmaterial)="" or isnull(patternmaterial) then patternmaterial ="NULL"
    if trim(patternmaterial)<>"NULL" then patternmaterial = pack(patternmaterial)
    
    if trim(patternowner)="" or isnull(patternowner) then patternowner ="NULL"
    if trim(patternowner)<>"NULL" then patternowner = pack(patternowner)
    
    if trim(patternavailability)="" or isnull(patternavailability) then patternavailability ="NULL"
    if trim(patternavailability)<>"NULL" then patternavailability = pack(patternavailability)

    
    sQuery = "Select ItemCode from INV_M_ItemOrgManufacturing where ItemCode ="& sItemCode &" and ClassificationCode = "& sClassCode &" and OrganisationCode ="& sOrgCode
    rstemp.Open sQuery,con
    if rstemp.EOF then
        sQuery = "Insert into INV_M_ItemOrgManufacturing (ItemCode,ClassificationCode,OrganisationCode,PartyCode,Alias,CasttingDrawNo,MachineDrawNo,Grade,MouldingProcess,NoofCavities,MatchPlateNo,ItemWeight,basevalue,ItemRate,exportrate,currency,PatternMaterial,PatternOwner,PatternAvailability)"&_
                 " values("&sItemCode&","& sClassCode&","& Pack(sOrgCode) &","& partycode &","& alias &","& cdn &","& mdn &","& grade & ","& modelprocess &","& noofcavities &","& plateno &","& itemweight &","& sbasevalue &","& sitemrate &","& sexportrate &","& scurrency &","& patternmaterial &","& patternowner &","& patternavailability &")"
    else
    
        sQuery = "Update INV_M_ItemOrgManufacturing set PartyCode="& partycode &",Alias="& alias &",CasttingDrawNo="& cdn &",MachineDrawNo="& mdn &",Grade="& grade & ",MouldingProcess="& modelprocess &",NoofCavities="& noofcavities &",MatchPlateNo="& plateno &",ItemWeight="& itemweight &","&_
                           "basevalue="& sbasevalue &",ItemRate="& sitemrate &",exportrate="& sexportrate &",currency="& scurrency &",PatternMaterial="& patternmaterial &",PatternOwner="& patternowner &",PatternAvailability="& patternavailability &""&_
                           " where ItemCode ="&sItemCode&" and ClassificationCode = "& sClassCode&" and OrganisationCode ="& Pack(sOrgCode) 
    end if
    rstemp.Close 
    
    Response.Write "<p>"& sQuery
    con.execute sQuery
    
    
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