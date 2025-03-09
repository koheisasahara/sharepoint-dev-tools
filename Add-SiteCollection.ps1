# Description: This script adds a new site collection to the SharePoint Online tenant.
# Usage: ./Add-SiteCollection.ps1

try{
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
    
    Write-Host "Title: $($web.Title)"
    Write-Host "URL: $($web.Url)"
    Write-Host "Description: $($web.Description)"
    Write-Host "Created: $($web.Created)"
    Write-Host "Last Modified: $($web.LastItemModifiedDate)"
    Write-Host "ID: $($web.Id)"
    Write-Host "Language: $($web.Language)"
    Write-Host "Server Relative URL: $($web.ServerRelativeUrl)"
}
catch{
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally{
    # Disconnect the context
    $context.Dispose()
}
