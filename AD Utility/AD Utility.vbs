Option Explicit

Dim objExcel, objWorkbook
Dim strDomain, strGroup, strGroup1, strUser
Dim intRow
Dim objGroup, objGroup1, objUser
Dim dictUsers, dictUsers1
Dim blnIsMember
Dim strTimestamp
Dim arrInGroup, arrNotInGroup, arrInGroup1, arrNotInGroup1
Dim intTotalUsers

Set objExcel = CreateObject("Excel.Application")

' Open the Excel file
Set objWorkbook = objExcel.Workbooks.Open("C:\AD Utility\AD Input.xlsx")

' Read the domain and group from the first worksheet
strDomain = objWorkbook.Worksheets("AD group").Cells(1, 1).Value
strGroup = objWorkbook.Worksheets("AD group").Cells(2, 1).Value
strGroup1 = objWorkbook.Worksheets("AD group").Cells(3, 1).Value

' Print domain and group being checked
WScript.Echo "Checking membership of the service accounts in the AD group as on " & Now
WScript.Echo "AD group: '" & strGroup & "'."
WSCript.Echo "Domain: '" & strDomain & "'."

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
intRow = 2
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
WScript.Echo " "
WScript.Echo "Users in the AD group:"
For Each strUser In arrInGroup
    WScript.Echo "  - " & strUser
Next

' Print users not in group
WScript.Echo " "
WScript.Echo "Users not in the AD group:"
For Each strUser In arrNotInGroup
    WScript.Echo "  - " & strUser
Next

' Print total number of users and the number of users in and not in the group
WScript.Echo " "
intTotalUsers = UBound(arrInGroup) + UBound(arrNotInGroup) + 2
'WScript.Echo intTotalUsers & " users were inputted."
'WScript.Echo (UBound(arrInGroup) + 1) & " users belong to the AD group."
'WScript.Echo (UBound(arrNotInGroup) + 1) & " users do not belong to the AD group."

WScript.Echo "Out of "& intTotalUsers &" accounts that were searched, "& (UBound(arrInGroup) + 1) &" accounts was(were) part of the AD group and "& (UBound(arrNotInGroup) + 1) &" was(were) not part of the AD group."
'WScript.Echo "Command Executed Successfully on: " & Now 
'WScript.Echo " "


' Print timestamp
'strTimestamp = Now
WScript.Echo " "

' Print domain and group being checked
WScript.Echo "Checking membership of the service accounts in the AD group as on " & Now
WScript.Echo "AD group: '" & strGroup1 & "'."
WSCript.Echo "Domain: '" & strDomain & "'."

Set objGroup1 = GetObject("WinNT://" & strDomain & "/" & strGroup1 & ",group")

Set dictUsers1 = CreateObject("Scripting.Dictionary")
For Each objUser In objGroup1.Members
    dictUsers1(objUser.Name) = True
Next

arrInGroup1 = Array()
arrNotInGroup1 = Array()

' Read the users from the second worksheet
intRow = 2
Do While objWorkbook.Worksheets("User List").Cells(intRow, 1).Value <> ""
    strUser = objWorkbook.Worksheets("User List").Cells(intRow, 1).Value
    blnIsMember = dictUsers1.Exists(strUser)
    If blnIsMember Then
        ReDim Preserve arrInGroup1(UBound(arrInGroup1) + 1)
        arrInGroup1(UBound(arrInGroup1)) = strUser
    Else
        ReDim Preserve arrNotInGroup1(UBound(arrNotInGroup1) + 1)
        arrNotInGroup1(UBound(arrNotInGroup1)) = strUser
    End If
    intRow = intRow + 1
Loop


' Print users in group
WScript.Echo " "
WScript.Echo "Users in the AD group:"
For Each strUser In arrInGroup1
    WScript.Echo "  - " & strUser
Next

' Print users not in group
WScript.Echo " "
WScript.Echo "Users not in the AD group:"
For Each strUser In arrNotInGroup1
    WScript.Echo "  - " & strUser
Next

' Print total number of users and the number of users in and not in the group
WScript.Echo " "
intTotalUsers = UBound(arrInGroup1) + UBound(arrNotInGroup1) + 2
'WScript.Echo intTotalUsers & " users were inputted."
'WScript.Echo (UBound(arrInGroup1) + 1) & " users belong to the AD group."
'WScript.Echo (UBound(arrNotInGroup1) + 1) & " users do not belong to the AD group."

WScript.Echo "Out of "& intTotalUsers &" accounts that were searched, "& (UBound(arrInGroup1) + 1) &" accounts was(were) part of the AD group and "& (UBound(arrNotInGroup1) + 1) &" was(were) not part of the AD group."
WScript.Echo "Command Executed Successfully on: " & Now 
WScript.Echo " "


' Clean up
objWorkbook.Close
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objGroup = Nothing
Set objUser = Nothing
Set dictUsers = Nothing
Set dictUsers1 = Nothing