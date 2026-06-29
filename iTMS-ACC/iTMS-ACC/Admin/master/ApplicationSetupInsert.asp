
<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ApplicationSetupInsert.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jun 30,2012
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

<!--#include file="../../include/DatabaseConnection.asp"-->
<!-- #include File="../../include/populate.asp" -->
<%
    Dim xmlDOM,objFSO
    Dim ndRoot,ndRow,ndDet
    Dim sQuery,sEntryNo,sAppCode,sRefCode
    Dim sMREntry,sIssueEntry,sReceiptAcc,sGRNEntry,sConEntry,sPOSEntry,sCRBEntry,sAutoAcc
    
    set xmlDOM = Server.CreateObject("Microsoft.XMLDOM")
    set objFSO = CreateObject("Scripting.FileSystemObject")
    con.begintrans
    if objFSO.FileExists(server.MapPath("../temp/ApplicationSetup_"&Session.SessionID&".xml")) then
        
        xmlDOM.load(server.MapPath("../temp/ApplicationSetup_"&Session.SessionID&".xml"))
        set ndRoot = xmlDOM.documentElement
       ' Response.ContentType = "text/xml"
       ' Response.Write ndRoot.xml
        if ndRoot.hasChildNodes() then
            for each ndRow in ndRoot.childNodes
                if ndRow.nodeName="Row" then
                    for each ndDet in ndRow.childNodes
                        if ndDet.nodeName="Det" then
                            sEntryNo = ndRow.getAttribute("EntryNo")
                            sAppCode = ndRow.getAttribute("AppCode")
                            sRefCode = ndRow.getAttribute("RefCode")
                    
                            sMREntry = ndDet.getAttribute ("MREntry")
                            sIssueEntry = ndDet.getAttribute("IssEntry")
                            sReceiptAcc = ndDet.getAttribute("RcptAcc")
                            sGRNEntry = ndDet.getAttribute("GRN")
                            sConEntry = ndDet.getAttribute("ConEntry")
                            sPOSEntry = ndDet.getAttribute("ManPOS")
                            sCRBEntry = ndDet.getAttribute("ComRcptBill")
                            sAutoAcc = ndDet.getAttribute("AutoAcc")
                            
                            sQuery = "Update APP_M_ApplicationSetup set AutomaticMREntry="& pack(sMREntry) &",AutomaticIssueEntry="& pack(sIssueEntry) &","
                            sQuery = sQuery & "AutomaticRcptAccounting="& pack(sReceiptAcc) &",AutomaticGatepassEntry="& pack(sGRNEntry) &",AutomaticConsumptionEntry="& pack(sConEntry) &","
                            sQuery = sQuery & "MandatoryPOS="& pack(sPOSEntry) &",CommonRcptBillEntry="& pack(sCRBEntry) &",AutomaticAccounting="& pack(sAutoAcc) &" where SetupEntryNo="& sEntryNo &" and ApplicationCode="& sAppCode &" and ReferenceCodeNo="& sRefCode
                            
                            Response.Write "<p>"& sQuery
                            con.execute sQuery
                            
                        end if 'if ndDet.nodeName="Det" then
                    next
                end if 'if ndRow.nodeName="Row" then
            next
        end if 'if ndRoot.hasChildNodes() then
        'objFSO.DeleteFile(server.MapPath("../temp/ApplicationSetup_"& Session.SessionID".xml"))
    end if 'if objFSO.FileExists(server.MapPath("../temp/ApplicationSetup_"&Session.SessionID&".xml")) then
'    con.rollbacktrans
 '   Response.End 
    Response.Clear 
    con.committrans
    
    Response.Redirect "ApplicationSetup.asp"
%>