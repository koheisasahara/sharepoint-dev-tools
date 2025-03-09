# Description: 
# Usage: ./Add-Lists.ps1

try{
    # Load the list csv file
    $listFile = Import-Csv -Path (Resolve-Path "$PSScriptRoot\data\list_input.csv")

    # Load the SharePoint Online CSOM Assemblies
    Add-Type -Path (Resolve-Path "$PSScriptRoot\modules\SharePointOnline.CSOM\1.0.10\Microsoft.SharePoint.Client.dll")
    Add-Type -Path (Resolve-Path "$PSScriptRoot\modules\SharePointOnline.CSOM\1.0.10\Microsoft.SharePoint.Client.Runtime.dll")
    
    $siteUrl = Read-Host "Enter the site collection URL"
    $username = Read-Host "Enter the username"
    $password = Read-Host "Enter the password" -AsSecureString
    
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
    
    $web = $context.Web
    $context.Load($web)
    $context.ExecuteQuery()
    
    try {
        foreach($list in $listFile){
            $listCreationInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
            $listCreationInfo.Title = $list.ListInternalName
            $listCreationInfo.TemplateType = $list.TemplateType
            $listCreationInfo.Description = $list.Description

            $newList = $web.Lists.Add($listCreationInfo)
            $context.Load($newList)
            $context.ExecuteQuery()
            
            Write-Host "List $($newList.ListInternalName) created successfully"
        }

        Start-Sleep -Seconds 5

        


    }
    catch {
        Write-Host $list.ListInternalName
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
catch{
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally{
    # Disconnect the context
    $context.Dispose()
}
