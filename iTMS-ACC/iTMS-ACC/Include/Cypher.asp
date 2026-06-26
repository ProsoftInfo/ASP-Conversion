<%
Private Function Encrypt(ByVal string)
	Dim x, i, tmp
	For i = 1 To Len( string )
		x = Mid( string, i, 1 )
		tmp = tmp & Chr( Asc( x ) + 6 )
	Next
	tmp = StrReverse( tmp )
	Encrypt = tmp
End Function

Private Function Decrypt(ByVal encryptedstring)
	Dim x, i, tmp
	encryptedstring = StrReverse( encryptedstring )
	For i = 1 To Len( encryptedstring )
		x = Mid( encryptedstring, i, 1 )
		tmp = tmp & Chr( Asc( x ) - 6 )
	Next
	Decrypt = tmp
End Function
%>