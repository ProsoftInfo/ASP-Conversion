<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XmlGetActTemp.asp
	'Module Name				:	Activity Role
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
    Dim objDOM,rsObj,rsTemp,rsAct,sQuery,ndRoot,ndTemplate,ndActivity,ndTemp
    Dim sActCode,sAppCode,sProcessCode,iTempCnt,iEligible
    
    set objDOM =Server.CreateObject("Microsoft.XMLDOM")
    set rsObj = Server.CreateObject("ADODB.Recordset")
    set rsAct = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    
    sAppCode = Request.QueryString("AppCode")
    sProcessCode =  Request.QueryString("ProcessCode")
    sActCode = Request.QueryString("ActCode")
    
    set ndRoot = objDOM.createElement("Root")
    objDOM.appendChild ndRoot
    
    iEligible = 0
    
    sQuery = "Select ApplicationCode,ProcessCode,ActivityCode,ActivityName,Status from Ms_ApplicationActivity where ApplicationCode = "& sAppCode &" and ProcessCode= "& sProcessCode &" and ActivityCode in ("& sActCode &")"
    rsAct.open sQuery,con
    if not rsAct.eof then
        do while not rsAct.eof
            iTempCnt = 0
            
            set ndActivity = objDOM.createElement("Activity")
                ndActivity.setAttribute "AppCode",rsAct(0)
                ndActivity.setAttribute "ProcessCode",rsAct(1)
                ndActivity.setAttribute "ActCode",rsAct(2)
                ndActivity.setAttribute "ActName",rsAct(3)
                ndActivity.setAttribute "Status",rsAct(4)
                ndRoot.appendChild ndActivity
                
                 sQuery = " Select ActivityTemplateNo,ActivityTemplateName,ProgramPath,Status from Ms_ApplicationActivityTemplates "&_
                         " where ApplicationCode = "& rsAct(0) &" and ProcessCode= "& rsAct(1) &" and ActivityCode =" & rsAct(2) &_
                         " and Cast(ActivityCode as varchar) +':'+ Cast(ActivityTemplateNo as varchar) in "&_
                         " (Select distinct Cast(ActivityCode as varchar)+':'+ Cast(ActivityTemplateNo as Varchar) from "&_
                         " MS_RoleActivity where PROCESSCODE = "& rsAct(1) &" AND APPLICATIONCODE = "& rsAct(0)&")"
                
                rsObj.Open sQuery,con
                if not rsObj.EOF then
                    do while not rsObj.EOF 
                        iTempCnt = iTempCnt + 1
                        
                        set ndTemp = objDOM.createElement("Template")
                        ndTemp.setAttribute "TempNo",rsObj(0)
                        ndTemp.setAttribute "TempName",rsObj(1)
                        ndTemp.setAttribute "Path",rsObj(2)
                        ndTemp.setAttribute "Status",rsObj(3)
                        ndTemp.setAttribute "Select","Y"
                        ndActivity.appendChild ndTemp
                        rsObj.MoveNext 
                    loop
                end if
                rsObj.Close 
                
                
                sQuery = " Select ActivityTemplateNo,ActivityTemplateName,ProgramPath,Status from Ms_ApplicationActivityTemplates "&_
                         " where ApplicationCode = "& rsAct(0) &" and ProcessCode= "& rsAct(1) &" and ActivityCode =" & rsAct(2) &_
                         " and Cast(ActivityCode as varchar) +':'+ Cast(ActivityTemplateNo as varchar) Not in "&_
                         " (Select distinct Cast(ActivityCode as varchar)+':'+ Cast(ActivityTemplateNo as Varchar) from "&_
                         " MS_RoleActivity where PROCESSCODE = "& rsAct(1) &" AND APPLICATIONCODE = "& rsAct(0)&")"
                
                rsObj.Open sQuery,con
                if not rsObj.EOF then
                    do while not rsObj.EOF 
                        iTempCnt = iTempCnt + 1
                        
                        set ndTemp = objDOM.createElement("Template")
                        ndTemp.setAttribute "TempNo",rsObj(0)
                        ndTemp.setAttribute "TempName",rsObj(1)
                        ndTemp.setAttribute "Path",rsObj(2)
                        ndTemp.setAttribute "Status",rsObj(3)
                        ndTemp.setAttribute "Select","N"
                        ndActivity.appendChild ndTemp
                        rsObj.MoveNext 
                    loop
                end if
                rsObj.close
                
                
                ndActivity.setAttribute "TempCnt",iTempCnt
                
                if iTempCnt > 1 then
                    iEligible = iEligible + 1
                end if
                
            rsAct.movenext
        loop
    end if
    rsAct.close
    
    if iEligible > 0 then
        ndRoot.setAttribute "Eligible","Y"
    else
        ndRoot.setAttribute "Eligible","N"
    end if
    
    
    Response.Clear
    Response.ContentType = "text/xml"
    Response.Write objDOM.xml
%>