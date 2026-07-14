
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmSpecsUpdatePop.asp
	'Module Name				:	Inventory (Item Modification)
	'Author Name				:	Ragavendarn
	'Created On					:	July 29,2011
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
    Dim iTypeHeadId,iAttId,sAttName,sAttValue,rsTemp
    Dim oDOM,RootNode,ndTempNode,ndAttNode,objFS
    Dim sItemType,iClass,iItmCode,sExp,iCnt,sSql
    
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    
    iClass= Request("hClassCode")
    iItmCode=Request("hItemCode")
    sItemType = Request("hItemType")
    
    Response.Write "<font color=red>"
    con.begintrans
    
    set oDOM = Server.CreateObject("Microsoft.XMLDOM")
    set objFS = Server.CreateObject("Scripting.FileSystemObject")
    
    if objFS.FileExists(Server.MapPath("../temp/Master/ItemSpecs"& Session.SessionID &".xml")) then
        oDOM.load(Server.MapPath("../temp/Master/ItemSpecs"& Session.SessionID &".xml"))
        
        
        set RootNode = oDOM.documentElement
        
        sExp = "//TypeHeader"
	    set ndTempNode = RootNode.selectNodes(sExp)
	    if ndTempNode.length>0 then
	        iTypeHeadId = ndTempNode.Item(0).Attributes.getNamedItem("ID").Value
	    end if
	    
	    sExp = "//TypeHeader/ATTRIBUTE"
	    set ndAttNode = RootNode.selectNodes(sExp)
	    Response.Write "<p>ndAttNode.length = "& ndAttNode.length
	    if ndAttNode.length>0 then
	        For iCnt = 0 to ndAttNode.length -1
	            iAttId = ndAttNode.Item(iCnt).Attributes.getNamedItem("ID").value
	            sAttName = ndAttNode.Item(iCnt).Attributes.getNamedItem("NAME").value
	            sAttValue = ndAttNode.Item(iCnt).Attributes.getNamedItem("VALUE").value
	            
	            sSql = "Select * from INV_M_ItemMasterAttributes where ClassificationCode ="& iClass &" and ItemCode="& iItmCode &" and HeaderID="& iTypeHeadId &" and ItemTypeAttributeID="& iAttId
	            With rsTemp
	                .CursorLocation = 3
	                .CursorType =3
	                .ActiveConnection = con
	                .Source = sSql
	                .Open 
	            End with
	            if rsTemp.EOF then
	                sSql = "Insert into INV_M_ItemMasterAttributes (ClassificationCode,ItemCode,"&_
	                       " HeaderID,ItemTypeAttributeID,ItemTypeAttributeName,AttributeValue)"&_
	                       " values("& iClass &","& iItmCode &","& iTypeHeadId &","& iAttId &","&_
	                       " "&pack(sAttName)&","& Pack(sAttValue) &")"
	                Response.Write "<p>"& sSql
	                con.execute sSql
	            else
	                sSql = " Update INV_M_ItemMasterAttributes set AttributeValue = "& Pack(sAttValue) &" where ClassificationCode ="& iClass &" and ItemCode="& iItmCode &" and HeaderID="& iTypeHeadId &" and ItemTypeAttributeID="& iAttId
	                Response.Write "<p>"& sSql
	                con.execute sSql
	            end if
	            rsTemp.Close 
	        Next
	    end if
    end if
	    
	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count - 1
			Response.Write con.Errors(iErrCounter) & "<BR>"
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
	
	Response.Redirect "ItmSpecsEditPop.asp?ItemCode="&iItmCode&"&ClassCode="& iClass &"&ItemType="&sItemType
%>
