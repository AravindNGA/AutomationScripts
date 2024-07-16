Option Explicit

Dim objExcel, objWorkbook
Dim strDomain, strGroup, strUser
Dim intRow
Dim objGroup, objUser
Dim dictUsers
Dim blnIsMember

Set objExcel = CreateObject("Excel.Application")

' Open the Excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\aravng\Desktop\ADTesting\AD Utility 3.0\Input.xlsx")

' Read the domain and group from the first worksheet
strDomain = objWorkbook.Worksheets("AD group").Cells(1, 1).Value
strGroup = objWorkbook.Worksheets("AD group").Cells(2, 1).Value

' Connect to the group
Set objGroup = GetObject("WinNT://" & strDomain & "/" & strGroup & ",group")

' Enumerate users in the group and load them into dictionary
Set dictUsers = CreateObject("Scripting.Dictionary")
For Each objUser In objGroup.Members
    dictUsers(objUser.Name) = True
Next

' Read the users from the second worksheet
intRow = 1
Do While objWorkbook.Worksheets("User List").Cells(intRow, 1).Value <> ""
    strUser = objWorkbook.Worksheets("User List").Cells(intRow, 1).Value
    blnIsMember = dictUsers.Exists(strUser)
    If blnIsMember Then
        WScript.Echo strUser & " is part of the AD group."
    Else
        WScript.Echo strUser & " is not part of the AD group."
    End If
    intRow = intRow + 1
Loop

' Clean up
objWorkbook.Close
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objGroup = Nothing
Set objUser = Nothing
Set dictUsers = Nothing