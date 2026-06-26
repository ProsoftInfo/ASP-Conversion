<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetAttrName.asp
%>
<!-- #include File="DatabaseConnection.asp" -->
<%	
    Dim rsTemp
    Dim sQuery,sOptionValue,sAttName,sTempOptValue
    
    sOptionValue = Request("AttID")
    
    sTempOptValue = replace(sOptionValue,":",",")
    
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    
    sQuery = "Select OptionName from INV_M_ItemTypeOptions where OptionValue in (" & sTempOptValue & ")"
    rsTemp.open sQuery,con
    if not rsTemp.eof then
        do while not rsTemp.eof
            sAttName = sAttName & ","& rsTemp(0)
            rsTemp.movenext
        loop    
    end if
    rsTemp.close
    
    response.write mid(sAttName,2)
    
%>
