function Get-UsersByGroup{
param(
[string]$groupName
)
    if( $vault -eq $null){
        echo 'Not connected to vault'
        return
    }
    try{
        $group = $vault.AdminService.GetGroupByName($groupName)
        $groupInfo = $vault.AdminService.GetGroupInfoByGroupId($group.Id)
        $groupInfo.Users | ForEach-Object { [array]$users += $_ }
        return $users
    }catch [Autodesk.Connectivity.WebServices.VaultServiceErrorException]{
        if ($_.Exception.Message -eq 160){
            echo 'Bad group name'
        } else {
            echo $_.Exception.Message
        }
    }
}