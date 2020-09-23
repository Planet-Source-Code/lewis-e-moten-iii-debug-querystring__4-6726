<div align="center">

## Debug QueryString


</div>

### Description

Just used for debugging querystring data. Creates an orderd list of field names and the values assigned to each one.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__4-6.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-debug-querystring__4-6726/archive/master.zip)

### API Declarations

Copyright (c) 2001, Lewis Moten. All rights reserved.


### Source Code

Response.Write QueryStringData()
Function QueryStringData()
	Dim llngMaxFieldIndex
	Dim llngFieldIndex
	Dim llngMaxValueIndex
	Dim llngValueIndex
	Dim lstrDebug
	' Count QueryString
	llngMaxFieldIndex = Request.QueryString.Count
	' Let user know if QueryString do not exist
	If llngMaxFieldIndex = 0 Then
		QueryStringData = "QueryString data is empty."
		Exit Function
	End If
	' Begin building a list of all QueryString
	lstrDebug = "<OL>"
	' Loop through each QueryString
	For llngFieldIndex = 1 To llngMaxFieldIndex
		lstrDebug = lstrDebug & "<LI>" & Server.HTMLEncode(Request.QueryString.Key(llngFieldIndex))
		' Count the values
		llngMaxValueIndex = Request.QueryString(llngFieldIndex).Count
		' If the Field doesn't have multiple values ...
		If llngMaxValueIndex = 1 Then
			lstrDebug = lstrDebug & " = "
			lstrDebug = lstrDebug & Server.HTMLEncode(Request.QueryString.Item(llngFieldIndex))
		' Else loop through each value
		Else
			lstrDebug = lstrDebug & "<OL>"
			For llngValueIndex = 1 to llngMaxValueIndex
				lstrDebug = lstrDebug & "<LI>"
				lstrDebug = lstrDebug & Server.HTMLEncode(Request.QueryString(llngFieldIndex)(llngValueIndex))
				lstrDebug = lstrDebug & "</LI>"
			Next
			lstrDebug = lstrDebug & "</OL>"
		End If
		lstrDebug = lstrDebug & "</LI>"
	Next
	lstrDebug = lstrDebug & "</OL>"
	' Return the data
	QueryStringData = lstrDebug
End Function

