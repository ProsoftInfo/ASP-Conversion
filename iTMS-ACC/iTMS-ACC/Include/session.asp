<%
	dim sessId
	sessId = session("loginid")
	if isempty(sessId) or isnull(sessId) or trim(sessId) = "" then
		Response.Clear
		Session.Abandon
		Response.Redirect "../../include/SessionExpired.asp"
	end if
%>
