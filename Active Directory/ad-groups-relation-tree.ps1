# Import the Active Directory module
Import-Module ActiveDirectory

# Function to recursively get group members and their subgroups
function Get-GroupMembers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [int]$Level = 0
    )

    # Retrieve group members
    $members = Get-ADGroupMember -Identity $GroupName

    # Output group name with indentation based on level
    Write-Host ("  " * $Level) -NoNewline
    if($Level -ne 0){
        Write-Host ("└─ ") -NoNewline
    }
    Write-Host $GroupName

    # Recursive call for each member that is a group
    foreach ($member in $members) {
        if ($member.objectClass -eq "group") {
            Get-GroupMembers -GroupName $member.Name -Level ($Level + 2)
        }
    }
}

# Get all groups in Active Directory
$groups = Get-ADGroup -Filter * | Select-Object -ExpandProperty Name

# Loop through each group
foreach ($group in $groups) {
    # Call the function to get group members and illustrate the relation
    Get-GroupMembers -GroupName $group
}