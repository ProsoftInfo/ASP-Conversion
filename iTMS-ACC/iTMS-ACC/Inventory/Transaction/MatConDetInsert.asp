
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MatConDetInsert.asp
	'Module Name				:	Inventory
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Jun 24,2011
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
    Dim oDOM,objFSO,rsTemp,ndRoot,ndItem,ndLotSerDet,ndChild
    Dim ndWCenter,ndMCenter,sWCCode,sMCCode
    Dim sQuery,sRemarks,sStoresUOM,sAttributeList,sOrgID,sEntryNo
    Dim sLotNo,sSerNo,sLotAcHead,ndSer
    Dim iConNo,iLineNo,iIssueEntryNo,iItemCode,iClassCode,iIssQty
    Dim iConQty,iSerQty,iTotConQty,iCreatedBy,iMCQty,iTotQtyRet
    Dim dIssueDate,bFlagCons,iItemCount
    Dim RNo,RItem,RQty,RByProduct,iSerNo,iQty
    
    set oDOM = Server.CreateObject("Microsoft.XMLDOM")
    set objFSO = CreateObject("Scripting.FileSystemObject")
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    
    iCreatedBy = Session("userid")
    
    if objFSO.FileExists(Server.MapPath("../temp/transaction/MatConsumption"&Session.SessionID&".xml")) then
        oDOM.load Server.MapPath("../temp/transaction/MatConsumption"&Session.SessionID&".xml")
        set ndRoot = oDOM.documentElement
    else
        Response.Write "XML File not found"
        Response.End  
    end if
    
    con.begintrans
    
   
    
    If ndRoot.hasChildNodes() Then
        sOrgID = ndRoot.getAttribute("OrgID")
        For Each ndItem In ndRoot.childNodes
            If ndItem.nodeName="ItemDet" then
                iItemCode = ndItem.getAttribute("Item")
                iClassCode = ndItem.getAttribute("Class")
                iIssueEntryNo = ndItem.getAttribute("IssEntryNo")
                sRemarks = ndItem.getAttribute("Remarks")
                iTotConQty = ndItem.getAttribute("Qty")
                sAttributeList = ndItem.getAttribute("AttributeList")
                if trim(iItemCode)="" or IsNull(iItemCode) then iItemCode = "NULL" 
                if trim(iClassCode)="" or IsNull(iClassCode) then iClassCode = "NULL" 
                
                
                bFlagCons = false
                
                
                sQuery = "Select ConsumptionNo from INV_T_MaterialConsumption where ItemCode ="& iItemCode &" and IssueEntryNo = "& iIssueEntryNo   
                Response.Write "<p>"& sQuery
                rsTemp.open sQuery,con
                if not rsTemp.eof then
                    iConNo = rsTemp(0)
                    bFlagCons=true
                end if
                rsTemp.close
                
                if bFlagCons=false then
                    sQuery = "Select isNull(Max(ConsumptionNo)+1,1) from INV_T_MaterialConsumption"
                    Response.Write "<p>"& sQuery
                    rsTemp.Open sQuery,con
                    if not rsTemp.EOF then
                        iConNo = rsTemp(0)
                    end if
                    rsTemp.Close 
                end if
                
                if trim(iItemCode)<>"NULL" then
                    sQuery = "Select StoresUOM from VWITEM where ItemCode = "& iItemCode &" and ClassificationCode ="& iClassCode
                    Response.Write "<p>"& sQuery
                    rsTemp.Open sQuery,con
                    if not rsTemp.EOF then
                        sStoresUOM = trim(rsTemp(0))
                    end if
                    rsTemp.Close 
                end if 'if trim(iItemCode)<>"" then
                
                sQuery = "Select isNull(Max(LineNumber)+1,1) from INV_T_MaterialConsumption where ConsumptionNo = "& iConNo 
                Response.Write "<p>"& sQuery
                rsTemp.Open sQuery,con
                if not rsTemp.EOF then
                    iLineNo = rsTemp(0)
                end if
                rsTemp.Close 
                if trim(sAttributeList)<>"" then sAttributeList = Pack(sAttributeList)
                if trim(sAttributeList)="" then sAttributeList = "NULL"
                if trim(sRemarks)<>"" then sRemarks = Pack(sRemarks)
                if Trim(sRemarks)="" then sRemarks = "NULL"
                
                if bFlagCons=false then
                    sQuery = "Insert into INV_T_MaterialConsumption (ConsumptionNo,IssueEntryNo,LineNumber,"&_
                             " OrganisationCode,ClassificationCode,ItemCode,QuantityConsumed,QuantityUOM,"&_
                             " ApplicationCode,Remark,EnteredOn,EnteredBy,AttributeList) values ("& iConNo &","&_
                             " "& iIssueEntryNo &","& iLineNo &",'"& sOrgID &"',"& iClassCode &","& iItemCode &","&_
                             " "& iTotConQty &",'"& sStoresUOM &"',4,"& sRemarks &",Convert(datetime,getDate(),103),"& iCreatedBy &","& sAttributeList &")"
                             
                    Response.Write "<p>"& sQuery
                    con.execute sQuery
                end if 'if bFlagCons=false then
                
                For each ndChild in ndItem.childNodes
                    If ndChild.nodeName="LotDet" Then
                        sLotNo = ndChild.getAttribute("LotNo")
                        sLotAcHead = ndChild.getAttribute("AccHead")
                        dIssueDate = ndChild.getAttribute("IssueDate")
                        iTotQtyRet = ndChild.getAttribute("QtyRet")
                        if trim(sLotNo)="" or IsNull(sLotNo) then sLotNo = "NULL"
                        if trim(sLotNo)<>"NULL" then sLotNo = Pack(sLotNo)
                        if trim(sLotAcHead)="" or IsNull(sLotAcHead) then sLotAcHead="NULL"
                        
                        sQuery = "Select count(distinct ItemCode) from INV_T_MaterialIssueDetails where IssueEntryNo = "& iIssueEntryNo &" Group By ItemCode"
                        rsTemp.open sQuery,con
                        if not rsTemp.eof then
                            iItemCount = rsTemp(0)
                        end if 
                        rsTemp.close
                        
                        if ndChild.hasChildNodes() then
                            For each ndLotSerDet in ndChild.childNodes
                                sSerNo = ndLotSerDet.getAttribute("SerNo")
                                iSerQty = ndLotSerDet.getAttribute("Qty")
                                
                                if cdbl(iSerQty)>0 then
                                
                                    sQuery = "Insert into INV_T_MaterialConsumptionDetail (ConsumptionNo,LineNumber,IssueEntryNo,"&_
                                             " IssueDate,LotNo,SerialNo,QuantityConsumed,ConsumptionACHead,AttributeList) values("&_
                                             " "& iConNo &","& iLineNo &","& iIssueEntryNo &",Convert(datetime,'"& dIssueDate &"',103),"&_
                                             " "& sLotNo &","& sSerNo &","& iSerQty &","& sLotAcHead &","&sAttributeList &")"
                                    Response.Write "<p>"& sQuery
                                    con.execute sQuery
                                    
                                    if trim(iItemCode)="NULL" and iItemCount = 1 then
                                        sQuery = "Update INV_T_MaterialIssueDetails set QuantityConsumed = "& iSerQty &" where IssueEntryNo = "& iIssueEntryNo &" and SerialNo = "& sSerNo 
                                    else
                                        sQuery = "Update INV_T_MaterialIssueDetails set QuantityConsumed = "& iSerQty &" where IssueEntryNo = "& iIssueEntryNo &" and SerialNo = "& sSerNo &"  and ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode
                                    end if
                                    Response.Write "<p>"& sQuery
                                    con.execute sQuery
                                end if 'if iSerQty>0 then
                                   
                            Next
                        else
                            if cdbl(iTotQtyRet)>0 then
                                 sQuery = "Insert into INV_T_MaterialConsumptionDetail (ConsumptionNo,LineNumber,IssueEntryNo,"&_
                                          " IssueDate,LotNo,SerialNo,QuantityConsumed,ConsumptionACHead,AttributeList) values("&_
                                          " "& iConNo &","& iLineNo &","& iIssueEntryNo &",Convert(datetime,'"& dIssueDate &"',103),"&_
                                          "NULL,NULL,"& iTotQtyRet &","& sLotAcHead &","&sAttributeList &")"
                                 Response.Write "<p>"& sQuery
                                 con.execute sQuery
                                      
                                if trim(iItemCode)="NULL" and iItemCount = 1 then
                                    sQuery = "Update INV_T_MaterialIssueDetails set QuantityConsumed = "& iTotQtyRet  &" where IssueEntryNo = "& iIssueEntryNo
                                else
                                    sQuery = "Update INV_T_MaterialIssueDetails set QuantityConsumed = "& iTotQtyRet  &" where IssueEntryNo = "& iIssueEntryNo &" and ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode
                                end if
                                Response.Write "<p>"& sQuery
                                con.execute sQuery
                            end if 'if iSerQty>0 then
                        end if
                    Elseif ndChild.nodeName="AddDet" Then
                        sEntryNo = 0
                        For each ndWCenter in ndChild.childNodes
                            sWCCode = ndWCenter.getAttribute("WCODE")
                            For each ndMCenter in ndWCenter.childNodes
                                sMCCode = ndMCenter.getAttribute("MCODE")
                                iMCQty = ndMCenter.getAttribute("QTY")
                                sEntryNo = sEntryNo + 1
                                
                                if sMCCode ="select" or sMCCode ="" or IsNull(sMCCode) then sMCCode ="NULL"
							    if sMCCode <>"NULL" then sMCCode = pack(sMCode)
							    if sWCCode<>"" then sWCCode = Pack(sWCode)
								    
							    sQuery = " Insert into INV_T_MaterialConsumptionIssueAddnDet (LineNumber,ConsumptionNo,EntryNo,"&_
							               " WorkCenterCode,MachineCenterCode,MixCode,QuantityIssued,QuantityConsumed) values "&_
							               " ("& iLineNo &","& iConNo &","& sEntryNo &","& sWCCode &","& sMCCode &",NULL,"& iMCQty  &","& iMCQty  &")"
						        Response.Write "<p>"& sQuery
                                con.execute sQuery
                                
                            Next
                        Next
                    ElseIf ndChild.nodeName="RcptItem" Then
                    
                        RNo = ndChild.getAttribute("No")
                        RItem = ndChild.getAttribute("Item")
                        RQty = ndChild.getAttribute("Qty")
                        RByProduct = ndChild.getAttribute("ByProduct")
                        sQuery = "Insert into INV_T_MaterialConsumptionOutput(ConsumptionNo,IssueEntryNo,LineNumber,OrganisationCode,ClassificationCode,"&_
                                 "ItemCode,OutputItemCode,OutputQuantity,ByProduct,AppRefNo,AppRefType)"&_
                                 " values("& iConNo &","& iIssueEntryNo &","& iLineNo &",'"& sOrgID &"'," & iClassCode &","& iItemCode &","& RItem &","& RQty &",'"&RByProduct&"',"& RNo&",39)"
                        Response.Write "<p>"& sQuery
                        con.execute sQuery
                        
                        
                        if ndChild.hasChildNodes() then
                            for each ndSer in ndChild.childNodes
                                if ndSer.nodeName="SerialDetails" then
                                    iSerNo = ndSer.getAttribute("SerNo")
                                    iQty = ndSer.getAttribute("Qty")
                                    sQuery = "Insert into INV_T_MaterialConsumptionOutputDet(ConsumptionNo,LineNumber,IssueEntryNo,LotNo,SerialNo,OutputQuantity,AttributeList)"&_
                                             " values ("& iConNo &","& iLineNo &","& iIssueEntryNo &",NULL,"& iSerNo &","& iQty &",NULL)"
                                    Response.Write "<p>"& sQuery
                                    con.execute sQuery
                                end if
                            next
                        end if
                    End if 'If ndChild.nodeName="LotDet" Then
                Next 'For each ndChild in ndItem.childNodes
            End if 'if ndItem.nodeName="ItemDet" then
        Next
    End If
    
    if con.Errors.count <> 0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
	else
	'	con.RollbackTrans
	'	Response.End
	   Response.Clear
	   con.CommitTrans
	   Response.Redirect "MATERIALCONSUMPTIONS.asp"
	end if 'if con.Errors.count <> 0 then
%>
