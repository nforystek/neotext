<%
'***********************************************************************************************************
'***********************************************************************************************************
'These functions are used to mimic transactional response in case of error
'on error resume next

dim fullResponse

Function WebResponse(strOutPut)
	'Builds the response during page execution
	fullResponse = fullResponse & strOutPut
end function

Function FinishWebResponse()
	'Sends response or error when page is finished
	if Err.number = 0 then
		Response.Write fullResponse
	else
		Response.Write "ERROR " & trim(cStr(Err.number)) & " ToDoster.asp " & Err.Description
	end if
	Response.End
end function
'***********************************************************************************************************
'***********************************************************************************************************


	Dim dbConnection
	Set dbConnection = CreateObject("ADODB.Connection")

	Function OpenConnection()
	    'On Error Resume Next
	    If dbConnection.State <> 0 Then dbConnection.Close
	    dbConnection.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("ToDoster.mdb") & ";DefaultDir=" & Server.MapPath(".") & ";"
		'Err.Clear
	    OpenConnection = (Not dbConnection.State = 0)
	End Function

	Function OpenRecordSet(ByRef OnRecord, ByVal SQLStr)
	    'On Error Resume Next
	    If OnRecord.State <> 0 Then OnRecord.Close
	    OnRecord.Open SQLStr, dbConnection, , 3
	    'Err.Clear
	    OpenRecordSet = (Not OnRecord.State = 0)
	End Function

	 Function CloseRecordSet(ByRef OnRecord)
	    If OnRecord.State <> 0 Then OnRecord.Close
	    Set OnRecord = Nothing
	End Function

	Function CloseConnection()
	    If dbConnection.State <> 0 Then dbConnection.Close
	    set dbConnection = nothing
	End Function


	Function URLEncode(encodeString)
	    Dim returnString
	    Dim currentChar

	    Dim i

	    For i = 1 To Len(encodeString)
	        currentChar = Mid(encodeString, i, 1)

	        If Asc(currentChar) = 32 Then
	            returnString = returnString + "+"
	            'character is a space so change it to a plus
	        Else
	            'character needs a hex conversion
	            If Len(Hex(Asc(currentChar))) = 1 Then
	                returnString = returnString + "%0" + Hex(Asc(currentChar))
	            ElseIf Len(Hex(Asc(currentChar))) = 2 Then
	                returnString = returnString + "%" + Hex(Asc(currentChar))
	            End If
	        End If
	    Next
	    'Return the url encoded string
	    URLEncode = returnString
	End Function


	dim strMethod, rs
		

	strMethod = lcase(Request("Method"))

	
	Select Case lcase(strMethod)

	Case "testconnection"
		
		WebResponse OpenConnection()
		
		CloseConnection		

	case "refreshlist"
		
		OpenConnection
		set rs = CreateObject("ADODB.RecordSet")
		        
		if lcase(request("Filter")) = "None" then
			OpenRecordSet rs, "SELECT * FROM Changes ORDER BY cDateTime;"
		else
			OpenRecordSet rs, "SELECT * FROM Changes WHERE NOT cStatus='" & Request("Filter") & "' ORDER BY cDateTime;"
		end if
    
		WebResponse "<List>" & vbCrLf
    
		If Not rs.EOF And Not rs.BOF Then
		    
		    rs.MoveFirst
		    Do
		        WebResponse vbTab & "<Entry>" & vbCrLf
		        WebResponse vbTab & vbTab & "<ID>" & rs("ID") & "</ID>" & vbCrLf
		        WebResponse vbTab & vbTab & "<c1>" & URLEncode(rs("cDateTime") & "") & "</c1>" & vbCrLf
		        WebResponse vbTab & vbTab & "<c2>" & URLEncode(rs("cProduct") & "") & "</c2>" & vbCrLf
		        WebResponse vbTab & vbTab & "<c3>" & URLEncode(rs("cType") & "") & "</c3>" & vbCrLf
		        WebResponse vbTab & vbTab & "<c4>" & URLEncode(rs("cComments") & "") & "</c4>" & vbCrLf
		        WebResponse vbTab & vbTab & "<c5>" & URLEncode(rs("cStatus") & "") & "</c5>" & vbCrLf
		        WebResponse vbTab & "</Entry>" & vbCrLf
		        
		    
		        rs.MoveNext
		    Loop Until rs.EOF Or rs.BOF
		    
		End If

		WebResponse "</List>" & vbCrLf
    
		CloseRecordSet rs
		CloseConnection
		
	case "insertrecord"
		OpenConnection
		set rs = CreateObject("ADODB.RecordSet")
    
		Dim tmpID
    
		tmpID = -Int((10000 * Rnd) + 1) & " tmp " & -Int((10000 * Rnd) + 1)
    
		OpenRecordSet rs, "INSERT INTO Changes (cDateTime, cProduct, cType, cComments, cStatus) VALUES ('" & tmpID & "','" & Request("cProduct") & "','" & Request("cType") & "','" & Request("cComments") & "','" & Request("cStatus") & "');"
    
		OpenRecordSet rs, "SELECT * FROM Changes WHERE cDateTime='" & tmpID & "';"
    
		tmpID = rs("ID")
    
		OpenRecordSet rs, "UPDATE Changes SET cDateTime='" & Request("cDateTime") & "' WHERE ID=" & rs("ID") & ";"

		WebResponse CLng(tmpID)

		CloseRecordSet rs
		CloseConnection

	case "updaterecord"
		OpenConnection
		set rs = CreateObject("ADODB.RecordSet")

		OpenRecordSet rs, "UPDATE Changes SET cDateTime='" & Request("cDateTime") & "', cProduct='" & Request("cProduct") & "', cType='" & Request("cType") & "', cComments='" & Request("cComments") & "', cStatus='" & Request("cStatus") & "' WHERE ID=" & Request("ID") & ";"

		CloseRecordSet rs
		CloseConnection
			
		
	case "deleterecord"
		OpenConnection
		set rs = CreateObject("ADODB.RecordSet")

	    OpenRecordSet rs, "DELETE * FROM Changes WHERE ID=" & Request("ID") & ";"

		CloseRecordSet rs
		CloseConnection

	
	End Select


'Call FinishWebResponse, if error,
'then the error is sent back,
'otherwise the built up response is sent.
FinishWebResponse




%>