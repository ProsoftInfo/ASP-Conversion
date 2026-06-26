<%
	'Program Name				:	clsDatabase.asp
	'Module Name				:
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 16, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="DatabaseConnection.asp"-->
<SCRIPT language="VBScript" runat="Server">
' —————————————————————————————————————————————————————————————————————————————
	Class clsDatabase
' —————————————————————————————————————————————————————————————————————————————
		' Define internal variables
		Private mbIsSQLServer			' determines if we are talking to an SQL server

		Private mnAbsolutePage			' determines current position of absolute page
		Private mnPageCount				' Number of pages returned by recordset
		Private mnPageSize				' determines amount of records are returned on each page

		Private moRs					' Database Recordset

		Private mnRecordsAffected		' Amount of records affected by SQL Query

' —————————————————————————————————————————————————————————————————————————————
		Private Sub Class_Initialize()

			' Initialize internal properties
			mnRecordsAffected	= -1
			mnAbsolutePage		= 1
			mnPageSize			= -1
			mnPageCount			= -1

			' Attempt to open database connection
			Call OpenDatabase()

		End Sub ' Class_Initialize()
' —————————————————————————————————————————————————————————————————————————————
		Private Sub Class_Terminate()

			' Close the database connection
			CloseDatabase()

		End Sub ' Class_Terminate()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get AbsolutePage()

			AbsolutePage = mnAbsolutePage

		End Property ' Get AbsolutePage()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let AbsolutePage(ByRef anAbsolutePage)

			mnAbsolutePage = anAbsolutePage

		End Property ' Let AbsolutePage()
' —————————————————————————————————————————————————————————————————————————————
		Public Sub CloseDatabase()

			con.Close
			Set con = Nothing

		End Sub ' CloseDatabase()
' —————————————————————————————————————————————————————————————————————————————
		Public Function OpenDatabase()
			' Create a recordset object
			Set moRs = Server.CreateObject("ADODB.Recordset")
			' Return Positive Results
			OpenDatabase = True
		End Function ' OpenDatabase()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get PageCount()

			PageCount = mnPageCount

		End Property ' Get PageCount()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get PageSize()

			PageSize = mnPageSize

		End Property ' Get PageSize()
' —————————————————————————————————————————————————————————————————————————————
		Public Property Let PageSize(ByRef anPageSize)

			mnPageSize = anPageSize

		End Property
' —————————————————————————————————————————————————————————————————————————————
		Public Property Get RecordsAffected()

			' Return the amount of records affected by last SQL query
			RecordsAffected = mnRecordsAffected

		End Property ' Get RecordsAffected()
' —————————————————————————————————————————————————————————————————————————————
		Public Function SetData(ByRef asSQL, ByRef avDataAry)

			Dim lnAbsolutePage		' Page number of record sets
			Dim lnPageSize			' Number of records on each page
			Dim lnPage				' Current page being worked with

			Dim lsFieldNames		' Comma-Space delimited list of field names
			Dim lsFieldTypes		' Comma-Space delmited list of field types

			' Setup default values (in case procedure failes)
			SetData				= False

			mnPageCount			= -1
			mnRecordsAffected	= -1

			' copy module variables to local ones
			lnAbsolutePage		= mnAbsolutePage
			lnPageSize			= mnPageSize

			' Reset module variables
			mnAbsolutePage		= 1
			mnPageSize			= -1

			If asSQL = "" Then
				Response.Write("<BR><FONT color=red>No SQL provided!!!</FONT><BR>")
				Exit Function
			End If

			' If the page size has been defined
			If Not lnPageSize	= -1 Then
				' Specify the Microsoft Client cursor
				moRs.CursorLocation = 3
				' Acquire data with ConnectionString
				Call moRs.Open(asSQL, con)
			' Else we are to retrieve all records
			Else
				' Acquire data with ConnectionString
				Call moRs.Open(asSQL, con)
			End If ' Not lnPageSize = -1
			' Grab the amount of records affected
			mnRecordsAffected = moRs.RecordCount

			' Return positive results
			SetData = True

			' If data was found
			If Not moRs.EOF Then

				' If the page size has not been defined
				If lnPageSize = -1 Then

					mnPageCount = 1

					' Pull data into an array
					avDataAry = moRS.GetRows

				' Else the page size has been defined
				Else

					' Define the page size
					moRs.PageSize = lnPageSize

					' Acquire the number of pages available
					mnPageCount = moRs.PageCount

					' Jump to the page that user wants to retrieve
					moRs.AbsolutePage = lnAbsolutePage

					' If records still exist
					If Not moRs.EOF Then

						' Grab records from the current page
						avDataAry = moRs.GetRows(lnPageSize, 0)

					End If ' Not moRs.EOF

				End If ' lnPageSize = -1

			End If ' Not moRS.EOF

			' If the recordset is open
			If moRs.State = 1 Then

				' Close Recordset
				moRS.Close

			End If ' moRs.State = ADODB.adStateOpen

		End Function ' SetData(ByRef asSQL, ByRef avDataAry)

	End Class
' —————————————————————————————————————————————————————————————————————————————
</SCRIPT>
