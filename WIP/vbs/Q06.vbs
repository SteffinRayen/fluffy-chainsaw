'Function button06()
	'Create a Script to connect to a Excel DB 
	'and fetch room rent for room  number #003.
	 
	'Table Info:
	 
	'Room Number   Room Rent
	'001           1000
	'002           6000
	'003           8000
	'004           7000
	 
	filePath = "C:\Users\Steffin Rayen\Desktop\GIT_WORKSPACE\fluffy-chainsaw\WIP\files\Q06.accdb"
	SQLquery = "Select * from Customer where CustomerID = 1"
	 
	'Creation of Connection Object
	Set objConnection = CreateObject("ADODB.Connection")
	 
	'Creation of Recordset Object
	Set objRecordSet = CreateObject("ADODB.Recordset")
	 
	'Connecting to Access DB
	objConnection.open ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="&filePath&";")
	 
	'Executing Query
	objRecordSet.open SQLquery,objConnection
	 
	'retrieving the result
	Msgbox objRecordSet.getstring 
	 
	'Release objects
	Set objRecordSet = nothing
	Set objConnection = nothing
'End Function