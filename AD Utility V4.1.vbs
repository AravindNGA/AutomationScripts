Option Explicit

Dim objExcel, objWorkbook
Dim strDomain, strGroup, strUser
Dim intRow
Dim objGroup, objUser
Dim dictUsers
Dim blnIsMember
Dim strTimestamp
Dim arrInGroup, arrNotInGroup
Dim intTotalUsers

Set objExcel = CreateObject("Excel.Application")

' Open the Excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\AD Utility 3.0\AD Input.xlsx")

' Read the domain and group from the first worksheet
strDomain = objWorkbook.Worksheets("AD group").Cells(1, 1).Value
strGroup = objWorkbook.Worksheets("AD group").Cells(2, 1).Value

' Print domain and group being checked
WScript.Echo "Checking group '" & strGroup & "' on domain '" & strDomain & "'..."

' Connect to the group
Set objGroup = GetObject("WinNT://" & strDomain & "/" & strGroup & ",group")

' Enumerate users in the group and load them into dictionary
Set dictUsers = CreateObject("Scripting.Dictionary")
For Each objUser In objGroup.Members
    dictUsers(objUser.Name) = True
Next

' Initialize arrays to hold users
arrInGroup = Array()
arrNotInGroup = Array()

' Read the users from the second worksheet
intRow = 1
Do While objWorkbook.Worksheets("User List").Cells(intRow, 1).Value <> ""
    strUser = objWorkbook.Worksheets("User List").Cells(intRow, 1).Value
    blnIsMember = dictUsers.Exists(strUser)
    If blnIsMember Then
        ReDim Preserve arrInGroup(UBound(arrInGroup) + 1)
        arrInGroup(UBound(arrInGroup)) = strUser
    Else
        ReDim Preserve arrNotInGroup(UBound(arrNotInGroup) + 1)
        arrNotInGroup(UBound(arrNotInGroup)) = strUser
    End If
    intRow = intRow + 1
Loop

' Print users in group
WScript.Echo "Users in the AD group:"
For Each strUser In arrInGroup
    WScript.Echo "  " & strUser
Next

' Print users not in group
WScript.Echo "Users not in the AD group:"
For Each strUser In arrNotInGroup
    WScript.Echo "  " & strUser
Next

' Print total number of users and the number of users in and not in the group
intTotalUsers = UBound(arrInGroup) + UBound(arrNotInGroup) + 2
WScript.Echo intTotalUsers & " users were inputted."
WScript.Echo (UBound(arrInGroup) + 1) & " users belong to the AD group."
WScript.Echo (UBound(arrNotInGroup) + 1) & " users do not belong to the AD group."

' Print timestamp
strTimestamp = Now
WScript.Echo "Checked at: " & strTimestamp

' Clean up
objWorkbook.Close
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objGroup = Nothing
Set objUser = Nothing
Set dictUsers = Nothing