<%Function getCurrentDate()
Dim rsGetDate,sQuery
set rsGetDate = Server.CreateObject("ADODB.RecordSet")
    sQuery = "Select Convert(varchar,getdate(),103) "
    rsGetDate.open sQuery,con
    if not rsGetDate.eof then
        getCurrentDate = rsGetDate(0)
    end if 
    rsGetDate.close
End Function
%>