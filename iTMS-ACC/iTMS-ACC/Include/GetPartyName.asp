<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetPartyName.asp
%>
<!--#include virtual="/Include/DatabaseConnection.asp"-->
<%	
    Dim rsTemp
    Dim sQuery,sPartyCode,sPartyName
    
    sPartyCode = Request("ParCode")
    
    
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    
    sQuery = "Select PartyName from APP_M_PartyMaster where PartyCode in (" & sPartyCode & ")"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        sPartyName = rsTemp(0)
    end if
    rsTemp.close
    
    response.write sPartyName
    
%>
