<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppActivityCreationAndAmendInsert.asp
	'Module Name				:	Admin (Activity Creation)
	'Author Name				:	UMAMAHESWARI S
	'Created On					:	December 14, 2003
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
dim dcrs,dcrs1,sSql,iInternalActivityID,objDOM,sType,nAppCode,sAppName,nProcessCode
dim sActName,sPrgPath,sStatus,nActCode,sUser,sItemTypeID,sUnit,sQuery
Dim nActTempNo,sActTempName,sActDesc,sActStatus
dim ndRoot,ndActivity,ndTemplate,sProgramType,sEmail,sSMS
con.beginTrans

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
objDOM.async = False
objDOM.load(request)

Set ndRoot = objDOM.documentElement

if ndRoot.hasChildNodes() then
    for each ndActivity in ndRoot.childNodes
        if ndActivity.nodeName="Activity" then
            
            sType = ndActivity.getAttribute("TYPE")
            nAppCode = ndActivity.getAttribute("APPCODE")
            sAppName = ndActivity.getAttribute("APPNAME")
            nProcessCode = ndActivity.getAttribute("PROCESSCODE")
            nActCode = ndActivity.getAttribute("ACTIVITYCODE")
            sActName = ndActivity.getAttribute("ACTIVITYNAME")
            sActStatus = ndActivity.getAttribute("STATUS")
            
            
                If sType = "ADD" Then
                    sQuery = "SELECT ISNULL(MAX(ACTIVITYCODE)+1,1) FROM MS_APPLICATIONACTIVITY WHERE APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & ""
                    dcrs.open sQuery,con
                    if not dcrs.EOF then
	                    iInternalActivityID = trim(dcrs(0))
                    else
	                    iInternalActivityID = "1"
                    end if
                    dcrs.Close
                	
                    sQuery = "SELECT ACTIVITYNAME FROM MS_APPLICATIONACTIVITY WHERE LOWER(ACTIVITYNAME) = " & Pack(lcase(sActName)) & " AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & ""
                     with dcrs
                        .cursorLocation = 3
                        .CursorType = 3
                        .source = sQuery
                        .ActiveConnection = con
                        .open
                    end with
                    if dcrs.EOF then
	                    sSql = "INSERT INTO MS_APPLICATIONACTIVITY(APPLICATIONCODE,PROCESSCODE,ACTIVITYCODE," &_
		                    "ACTIVITYNAME,PROGAMPATH,Status) VALUES " &_ 
		                    "(" & nAppCode & "," & nProcessCode & "," & iInternalActivityID & "," & Pack(sActName) & ", " &_
		                    " " & Pack(sPrgPath) & ",'A')"
	                    'Response.Write sSql & "<BR>"
	                    con.Execute sSql
                    end if
                    dcrs.Close
                    nActCode = iInternalActivityID
                Else
                    sQuery = "SELECT ACTIVITYNAME FROM MS_APPLICATIONACTIVITY WHERE LOWER(ACTIVITYNAME) = " & Pack(lcase(sActName)) & " AND APPLICATIONCODE = " & nAppCode & " AND PROCESSCODE = " & nProcessCode & " AND ACTIVITYCODE <> " & nActCode
                    with dcrs
                        .cursorLocation = 3
                        .CursorType = 3
                        .source = sQuery
                        .ActiveConnection = con
                        .open
                    end with
                    if dcrs.EOF then
                        sSql =  "UPDATE MS_APPLICATIONACTIVITY SET ACTIVITYNAME=" & Pack(sActName) & ",PROGAMPATH=" & Pack(sPrgPath) & ",Status = '"& sStatus &"'  WHERE APPLICATIONCODE=" & nAppCode & " AND PROCESSCODE=" & nProcessCode & " AND ACTIVITYCODE=" & nActCode &" "
                        'Response.Write sSql & "<BR>"
                        con.Execute sSql
                    End IF
                    dcrs.Close 
                End If	'If sType = "ADD" Then


            for each ndTemplate in ndActivity.childNodes
                if ndTemplate.nodeName="Template" then
                    nActTempNo = ndTemplate.getattribute("No")
                    sActTempName = ndTemplate.getattribute("Name")
                    sActDesc =  ndTemplate.getattribute("Description")
                    sPrgPath = ndTemplate.getattribute("ProgramPath")
                    sStatus = ndTemplate.getattribute("Status")
                    sProgramType = ndTemplate.getattribute("FileType")
                    sEmail = ndTemplate.getAttribute("EMAIL")
                    sSMS = ndTemplate.getAttribute("SMS")
                    
                    if trim(sActTempName)="" or IsNull(sActTempName) then sActTempName= "NULL"
                    if trim(sActTempName)<>"NULL" then sActTempName=pack(sActTempName)
                    
                    if trim(sActDesc)="" or IsNull(sActDesc) then sActDesc= "NULL"
                    if trim(sActDesc)<>"NULL" then sActDesc=pack(sActDesc)
                    
                    if trim(sPrgPath)="" or IsNull(sPrgPath) then sPrgPath= "NULL"
                    if trim(sPrgPath)<>"NULL" then sPrgPath=pack(sPrgPath)
                    if Trim(sEmail)="" then sEmail="N" 
                    if Trim(sSMS) ="" then sSMS = "N"
                    
                    sStatus = pack(sStatus)
                    
                    if trim(nActTempNo)<>"" then
                        sQuery = "Select count(*) from Ms_ApplicationActivityTemplates where ApplicationCode = "& nAppCode &" and ProcessCode= "& nProcessCode &" and ActivityCode = "& nActCode &" and ActivityTemplateNo = "& nActTempNo
                         with dcrs
                            .cursorLocation = 3
                            .CursorType = 3
                            .source = sQuery
                            .ActiveConnection = con
                            .open
                        end with
                        if not dcrs.eof then
                        
                            if dcrs(0) = 1 then
                                 if trim(sActStatus)<>"" then
                                     sSql =  "UPDATE MS_APPLICATIONACTIVITY SET Status = '"& sActStatus &"'  WHERE APPLICATIONCODE=" & nAppCode & " AND PROCESSCODE=" & nProcessCode & " AND ACTIVITYCODE=" & nActCode &" "
                                     con.Execute sSql
                                 end if 
                            end if
                        
                            sSql = "Update Ms_ApplicationActivityTemplates set ActivityTemplateName="& sActTempName &",ProgramPath="& sPrgPath &",TemplateDescription="& sActDesc &",Status="& sStatus &",ProgramType="& pack(sProgramType) &",SendAsEmail='"& sEmail &"',SendAsSMS='"& sSMS &"' where ApplicationCode="& nAppCode &" and ProcessCode="& nProcessCode &" and ActivityCode ="& nActCode &" and ActivityTemplateNo ="& nActTempNo
                            con.execute sSql
                        end if
                        dcrs.close
                    else
                        sQuery = "Select IsNull(Max(ActivityTemplateNo),0) + 1 from Ms_ApplicationActivityTemplates where ApplicationCode = "& nAppCode & " and ProcessCode= "& nProcessCode  &" and ActivityCode = "& nActCode    
                        dcrs.open sQuery,con
                        if not dcrs.eof then
                            nActTempNo = dcrs(0)
                        end if
                        dcrs.close 
                        if trim(sActTempName)<>"NULL" then
                            sQuery = "Insert into Ms_ApplicationActivityTemplates (ApplicationCode,ProcessCode,ActivityCode,"&_
                                     "ActivityTemplateNo,ActivityTemplateName,ProgramPath,TemplateDescription,Status,ProgramType,SendAsEmail,SendAsSMS) "&_
                                     " values("& nAppCode &","& nProcessCode &","& nActCode &","& nActTempNo &","& sActTempName &","&_
                                     ""& sPrgPath &","& sActDesc &","& sStatus &","& pack(sProgramType) &",'"& sEmail &"','"& sSMS &"')"
                                     
                            con.execute sQuery
                        else
                            if trim(sActStatus)<>"" then
                                 sSql =  "UPDATE MS_APPLICATIONACTIVITY SET Status = '"& sActStatus &"'  WHERE APPLICATIONCODE=" & nAppCode & " AND PROCESSCODE=" & nProcessCode & " AND ACTIVITYCODE=" & nActCode &" "
                                 con.Execute sSql
                             end if 
                        end if
                        
                    end if
                end if
            next
            Response.write "ActNo:"&nActCode
        end if
    next
end if

if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
