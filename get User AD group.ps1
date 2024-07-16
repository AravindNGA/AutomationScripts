$username = "d0a08yk"
$domain = "HOMEOFFICE.WAL-MART.COM"

$outputFile = "C:\Users\vn56my4\Desktop\Getting Membership\$username-Membership.txt"
$counter = 0
$currenDateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"

$user = [ADSI]"WinNT://$domain/$username,user"
$groups = $user.Groups()

Clear-Content -Path $outputFile -ErrorAction SilentlyContinue
New-Item -Path $outputFile -ItemType File -Force
Add-Content -Path $outputFile -Value " "
Write-output " "
Write-output "Membership for user: '$username' in the domain '$domain' as on $currenDateTime"
Add-Content -Path $outputFile -Value "Membership for user: '$username' in the domain '$domain' as on $currenDateTime"
Add-Content -Path $outputFile -Value " "
Write-output " " 
foreach ($group in $groups) {
    $counter ++
    $groupName = $group.GetType().InvokeMember("Name", 'GetProperty', $null, $group, $null)
    Add-Content -Path $outputFile -Value "$counter. $groupName"
    Write-output "$counter. $groupName"
    
}
Add-Content -Path $outputFile -Value " "
Add-Content -Path $outputFile -Value "Total number of groups: $counter"
Write-output " " 
Write-output "Total number of groups: $counter"