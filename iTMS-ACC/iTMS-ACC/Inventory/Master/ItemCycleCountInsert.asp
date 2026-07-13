
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'---------------------------
	'Note : this program will construct XML for updation.
	' updation will be taking care by Store Procedure
	'---------------------------
	'Program Name				:	ItemCycleCountInsert.asp
	'Module Name				:	Inventory (CycleCounting)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 25,2013
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/NoSeries.asp" -->
<!-- #include File="../../include/NoSeriesCommonFunctions.asp"-->
<%
    Dim rsObj,oDOM,objfs
    Dim ndRoot,ndCycle
    Dim sQuery,sOrgCode,sCCDate,sArrFinPeriod,sFinFrom,sFinTo,sMode
    Dim iCreatedBy,Item,iClass,CStock,CValue,CCQty,CCVal
    Dim iCycleCountEntryNo
    Dim bFlag 
    
    set rsObj = Server.CreateObject("ADODB.Recordset")
	set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	set objfs = Server.CreateObject("Scripting.FileSystemObject")

	iCreatedBy = getUserid
	sArrFinPeriod = split(Session("FinPeriod"),":")
	sFinFrom = "01/04/"& sArrFinPeriod(0)
	sFinTo = "31/03/"& sArrFinPeriod(1)
	
	if objfs.FileExists(Server.MapPath("../temp/master/Inv_CycleCount_"&Session.SessionID&".xml")) then
	
	    oDOM.async = false
    	
	    oDOM.load Server.MapPath("../temp/Master/Inv_CycleCount_"&session.sessionId&".xml")
    	
	    con.beginTrans
	    set ndRoot = oDOM.documentElement
	    iCycleCountEntryNo = ndRoot.getAttribute("CycleCountEntryNo")
	    sOrgCode = ndRoot.getAttribute("OrgCode")
	    sCCDate = ndRoot.getAttribute("CCDate")
	    sMode = ndRoot.getAttribute("Mode") ' N for New ,E for Edit
	    
	    if trim(sMode)="" or IsNull(sMode) then sMode="N"
	    
	    if trim(sMode)="N" then
	        sQuery = "Select IsNull(Max(CycleCountEntryNo),0)+1 from Inv_T_ItemCycleCount"
            Response.write "<p>"&sQuery
            rsObj.open sQuery,con
            if not rsObj.eof then
                iCycleCountEntryNo = rsObj(0)
            end if
            rsObj.close
        end if 'if trim(sMode)="N" then
        
        if trim(sMode)="N" then
            sQuery = "Insert into INV_T_ItemCycleCount (CycleCountEntryNo,OrganisationCode,FinancialYearFrom,"
            sQuery = sQuery & " FinancialYearTo,CycleCountDate,CycleCountDoneBy,CycleCountDoneOn)"
            sQuery = sQuery & " values("& iCycleCountEntryNo &",'"& sOrgCode &"',Convert(datetime,'"& sFinFrom &"',103),Convert(datetime,'"& sFinTo &"',103),"
            sQuery = sQuery & "Convert(datetime,'"& sCCDate &"',103),"&iCreatedBy &",Convert(datetime,'"& date() &"',103))"
            Response.write "<p>"& sQuery
            con.execute sQuery
        else
            sQuery = "Update INV_T_ItemCycleCount set CycleCountDoneBy="& iCreatedBy &",CycleCountDate=Convert(datetime,'"& sCCDate &"',103)"&_
                     " where CycleCountEntryNo ="& iCycleCountEntryNo
            Response.write "<p>"&sQuery
            con.execute sQuery
        end if
        
	    sQuery = "Delete from INV_T_ItemCycleCountHistory where CycleCountEntryNo ="& iCycleCountEntryNo
	    Response.write "<p>"&sQuery
	    con.execute sQuery
	    
	    if ndRoot.hasChildNodes() then
	        for each ndCycle in ndRoot.childNodes
	            Item = ndCycle.getAttribute("ItemCode")
	            iClass = ndCycle.getAttribute("ClassCode")
	            CStock = ndCycle.getAttribute("CStock")
	            CValue = ndCycle.getAttribute("CValue")
	            CCQty = ndCycle.getAttribute("CCQty")
	            
                sQuery = "Insert into INV_T_ItemCycleCountHistory (CycleCountEntryNo,ClassificationCode,ItemCode,FinancialYearFrom,"&_
                         "FinancialYearTo,CycleCountDate,CycleCountStock,CurrentStock,CurrentValue)"&_
                         "values("& iCycleCountEntryNo&","& iClass &","& Item &",Convert(datetime,'"& sFinFrom &"',103),Convert(datetime,'"& sFinTo &"',103),"&_
                         "Convert(datetime,'"& sCCDate &"',103),"& CCQty &","& CStock &","& CValue &")"
                Response.write "<p>"&sQuery
                con.execute sQuery
	        next
	    end if 'if ndRoot.hasChildNodes() then
	    
	    if con.Errors.count <> 0 then
		    dim iErrCounter
		    con.RollbackTrans
		    for iErrCounter=0 to con.Errors.count - 1
			    Response.Write con.Errors(iErrCounter) & "<BR>"
		    next
	    else
		   ' con.rollbacktrans
		   ' Response.end
	        Response.clear
	        con.CommitTrans
	        if objfs.FileExists(Server.MapPath("../temp/master/Inv_CycleCount_"&Session.SessionID&".xml")) then
			    objfs.DeleteFile server.MapPath("../temp/master/Inv_CycleCount_"&Session.SessionID&".xml")
		    end if
		    if objfs.FileExists(Server.MapPath("../temp/master/ItemCycleCount"&Session.SessionID&".xml")) then
		        objfs.DeleteFile server.MapPath("../temp/master/ItemCycleCount"&Session.SessionID&".xml")
		    end if
	    end if

	    con.close
	    set con = nothing
	    Response.redirect "ItemCycleCountGrid.asp"
	end if 'if objfs.FileExists(Server.MapPath("../temp/master/Inv_CycleCount_"&Session.SessionID&".xml")) then
	
	

%>
