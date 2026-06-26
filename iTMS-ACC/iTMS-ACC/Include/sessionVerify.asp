<%
	Dim sSessId
	sSessId = session("loginid")
	If isempty(sSessId) Or IsNull(sSessId) Or trim(sSessId) = "" Then
		Response.Clear
		Session.Abandon
		Response.Redirect "../../include/SessionExpired.asp"
	End If
%>
