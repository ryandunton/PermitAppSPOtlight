<#
.SYNOPSIS
    This script manages site permissions for App Registrations in SharePoint Online.
.DESCRIPTION
    This script is designed to manage the permissions of an App Registration on a SharePoint
    Online site collection. It uses the Microsoft Graph API to add or delete permissions.
    The script requires an OAuth2 token for authentication and the URL of the SharePoint site.
    It retrieves the site ID and updates the permissions accordingly.
.EXAMPLE
    # Check site for existing app permissions on a site
    PS C:\> .\Invoke-PermitAppSPOtlight.ps1

    # Check site for existing app permissions and provide everything in pipeline
    # Note: -ConnectionSecret expects a SecureString
    PS C:\> .\Invoke-PermitAppSPOtlight.ps1 -TenantId "c93d1a61-8c45-44f1-9484-205ce56e7ac4" `
                                            -ConnectionAppId "9e941cb1-4832-4231-a867-a9321171ca7f" `
                                            -ConnectionSecret $Secret `
                                            -SiteUrl "https://contosodev.sharepoint.com/sites/sitecollection/"

    # Give an app registration permissions to a site and provide everything in pipeline
    PS C:\> .\Invoke-PermitAppSPOtlight.ps1 -TenantId "c93d1a61-8c45-44f1-9484-205ce56e7ac4" `
                                            -ConnectionAppId "9e941cb1-4832-4231-a867-a9321171ca7f" `
                                            -SiteUrl "https://contosodev.sharepoint.com/sites/sitecollection/" `
                                            -AppId '43390439-c4be-4a78-a6f6-b98cca90a181' `
                                            -DisplayName 'My App' `
                                            -Add

    # Delete existing app permissions from a site
    PS C:\> .\Invoke-PermitAppSPOtlight.ps1 -TenantId "c93d1a61-8c45-44f1-9484-205ce56e7ac4" `
                                            -ConnectionAppId "9e941cb1-4832-4231-a867-a9321171ca7f" `
                                            -SiteUrl "https://contoso.sharepoint.com/sites/sitecollection/" `
                                            -Delete
.NOTES
    Version: 20231114.01
    Author:  Ryan Dunton https://github.com/ryandunton
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $TenantId,    
    [Parameter()]
    [string]
    $ConnectionAppId,
    [Parameter()]
    [securestring]
    $ConnectionSecret,
    [Parameter()]
    [string]
    $SiteUrl,
    [Parameter()]
    [string]
    $AppId,
    [Parameter()]
    [string]
    $DisplayName,
    [Parameter()]
    [string]
    $Role,
    [Parameter()]
    [switch]
    $Add,
    [Parameter()]
    [switch]
    $Delete
)

begin {
    
}

process {
    function Get-BearerToken {
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $TenantId,    
            [Parameter()]
            [string]
            $ConnectionAppId,
            [Parameter()]
            [securestring]
            $ConnectionSecret
        )
        
        begin {
            if (($null -eq $TenantId) -or ('' -eq $TenantId)) {$TenantId = Read-Host "[*] What is the tenant id?"}
            if (($null -eq $ConnectionAppId) -or ('' -eq $ConnectionAppId)) {$ConnectionAppId = Read-host "[*] What is the App Id you are connecting to the tenant with?"}
            if (($null -eq $ConnectionSecret) -or ('' -eq $ConnectionSecret)) {$ConnectionSecret = Read-Host "[*] What is your App Secret?" -AsSecureString}
            $TmpConnectionSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ConnectionSecret))
            Write-Host "[+] Connecting to $TenantId" -ForegroundColor Green
            $TmpHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $TmpHeaders.Add("Content-Type", "application/x-www-form-urlencoded")
            $TmpBody = "grant_type=client_credentials&client_id=$ConnectionAppID&client_secret=$TmpConnectionSecret&resource=https%3A%2F%2Fgraph.microsoft.com"
        }
        
        process {
            Write-Host "[-] getting bearer token... " -NoNewline
            try {
                $TmpResponse = Invoke-RestMethod "https://login.microsoftonline.com/$TenantId/oauth2/token" -Method 'POST' -Headers $TmpHeaders -Body $TmpBody
                $TmpBearerToken = $TmpResponse.access_token
                Write-Host "Success!" -ForegroundColor Green
            }
            catch {
                Write-Host "Totes Fail!" -ForegroundColor Red
                Write-Host "[-] $($Error[0])"
            }
        }
        
        end {
            return $TmpBearerToken
        }
    }
    function Get-SPOSiteID {
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $SiteUrl,
            [Parameter()]
            [string]
            $BearerToken
        )
        
        begin {
            Write-Host "[+] Getting SiteID" -ForegroundColor Green
            if (($null -eq $SiteUrl) -or ('' -eq $SiteUrl)) {$SiteUrl = Read-Host "[*] What is the SharePoint site path? (example https://contoso.sharepoint.com/sites/sitename/)"}
            $SiteUrl = $SiteUrl.Replace('https://','').Replace('/sites',':/sites')
            $TmpHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $TmpHeaders.Add("Content-Type", "application/json")
            $TmpHeaders.Add("Authorization", "Bearer $BearerToken")
        }
        
        process {
            Write-Host "[-] connecting to $SiteUrl " -NoNewline
            try {
                $TmpResponse = Invoke-RestMethod "https://graph.microsoft.com/v1.0/sites/$SiteUrl" -Method 'GET' -Headers $TmpHeaders
                $SiteId = $TmpResponse.id.split(',')[1]
                Write-Host "success!" -ForegroundColor Green
            }
            catch {
                Write-Host "Totes Fail!" -ForegroundColor Red
                Write-Host "[-] $($Error[0])"
                exit
            }
        }
        
        end {
            return $SiteId
        }
    }
    function Show-Banner {
        Write-Host -ForegroundColor Blue "
        __________                     .__  __    _____                 
        \______   \ ___________  _____ |___/  |_ /  _  \ ______ ______  
        |     ____/ __ \_  __ \/     \|  \   __/  /_\  \\____ \\____ \ 
        |    |   \  ___/|  | \|  Y Y  |  ||  |/    |    |  |_> |  |_> >
        |____|    \___  |__|  |__|_|  |__||__|\____|__  |   __/|   __/ 
        __________________________ \/__  .__  .__    \/|.__   |___    
        /   _____\______   \_____  \_/  |_|  | |__| ____ |  |___/  |_  
        \_____  \ |     ___//   |   \   __|  | |  |/ ___\|  |  \   __\ 
        /        \|    |   /    |    |  | |  |_|  / /_/  |   Y  |  |   
        /_______  /|____|   \_______  |__| |____|__\___  /|___|  |__|   
                \/                  \/            /_____/      \/     "
        Write-Host "             by https://github.com/ryandunton/PermitAppSPOtlight" -ForegroundColor White -NoNewline
        Write-Host -ForegroundColor Red "
        ------------------------------------------------------------------
        |" -NoNewline
        Write-Host "    `"Unlocking Seamless Access and Elevating Collaboration!`"" -NoNewline
        Write-Host -ForegroundColor Red "    |
        |" -NoNewline
        Write-Host "  *Now comes with 10x more elevating and even more seamlessness" -NoNewline
        Write-Host -ForegroundColor Red " |
        ------------------------------------------------------------------
        "
    }
    function Get-SPOSitePermissions {
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $SiteId,
            [Parameter()]
            [string]
            $BearerToken
        )
        
        begin {
            Write-Host "[+] Getting Site Permissions" -ForegroundColor Green
            $TmpHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $TmpHeaders.Add("Content-Type", "application/json")
            $TmpHeaders.Add("Authorization", "Bearer $BearerToken")
        }
        
        process {
            Write-Host "[-] connecting to graph... " -NoNewline
            try {
                $TmpResponse = Invoke-RestMethod "https://graph.microsoft.com/beta/sites/$SiteID/permissions" -Method 'GET' -Headers $TmpHeaders
                #$TmpResponse = $TmpResponse | ConvertTo-Json -Depth 5
                Write-Host "success!" -ForegroundColor Green
            }
            catch {
                Write-Host "Totes Fail!" -ForegroundColor Red
                Write-Host "[-] $($Error[0])"
            }
        }
        
        end {
            [string]$TmpPerms = "[-] $(($TmpResponse.value.grantedToIdentities.application | %{"$($_.displayname) ($($_.id))"}) -join ' | ')"
            if (($null -eq $TmpPerms) -or ('' -eq $TmpPerms) -or ('[-]  ()' -eq $TmpPerms)) {$TmpPerms = "[-] Empty"}
            return $TmpPerms
        }
    }
    function Remove-SPOSitePermissions {
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $SiteId,
            [Parameter()]
            [string]
            $BearerToken
        )
        
        begin {
            # Write-Host "[+] Getting Site Permissions" -Foregroundcolor Green
            $TmpGetHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $TmpGetHeaders.Add("Content-Type", "application/json")
            $TmpGetHeaders.Add("Authorization", "Bearer $BearerToken")
            $TmpDelHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $TmpDelHeaders.Add("Content-Type", "application/json")
            $TmpDelHeaders.Add("Authorization", "Bearer $BearerToken")
        }
        
        process {
            Write-Host "[-] connecting to graph... " -NoNewline
            try {
                $TmpResponse = Invoke-RestMethod "https://graph.microsoft.com/beta/sites/$SiteID/permissions" -Method 'GET' -Headers $TmpGetHeaders
                $TmpResponse | ConvertTo-Json -Depth 5 | Out-Null
                Write-Host "success!" -ForegroundColor Green
            }
            catch {
                Write-Host "Totes Fail!" -ForegroundColor Red
                Write-Host "[-] $($Error[0])"
            }
            if ($TmpResponse.Value.Count -gt 0) {
                # Extract dynamic values from JSON
                [array]$DynamicAppList = $TmpResponse.value | ForEach-Object {
                    $DynamicAppDisplayName = $_.grantedtoidentities.application.displayName
                    $DynamicAppId = $_.id
                    "$DynamicAppDisplayName ($DynamicAppId)"
                }
                # Set initial selection index
                $SelectedValueIndex = 0
                Write-Host "[+] Select the application you wish to remove" -ForegroundColor Green
                Write-Host ""
                Write-Host ""
                # Display the menu and handle user input
                do {
                    # Move cursor to top of menu area
                    [Console]::SetCursorPosition(0, [Console]::CursorTop - $DynamicAppList.Count)
                    for ($i = 0; $i -lt $DynamicAppList.Count; $i++) {
                        if ($i -eq $SelectedValueIndex) {
                            Write-Host "[>] $($DynamicAppList[$i])" -NoNewline
                        } else {
                            Write-Host "[ ] $($DynamicAppList[$i])" -NoNewline
                        }
                        #Clear any extra characters from previous lines
                        $SpacesToClear = [Math]::Max(0, ($DynamicAppList[0].Length - $DynamicAppList[$i].Length))
                        Write-Host (" " * $SpacesToClear) -NoNewline
                        Write-Host ""
                    }
                    # Get user input
                    $KeyInfo = $Host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown')
                    # Process arrow key input
                    switch ($KeyInfo.VirtualKeyCode) {
                        38 {  # Up arrow
                            $SelectedValueIndex = [Math]::Max(0, $SelectedValueIndex - 1)
                        }
                        40 {  # Down arrow
                            $SelectedValueIndex = [Math]::Min($DynamicAppList.Count - 1, $SelectedValueIndex + 1)
                        }
                    }
                } while ($KeyInfo.VirtualKeyCode -ne 13)  # Enter key

                $SelectedValue = $DynamicAppList[$SelectedValueIndex]
                $ApplicationId = $SelectedValue.Substring($SelectedValue.IndexOf("(")+1, $SelectedValue.IndexOf(")")-$SelectedValue.IndexOf("(")-1)

                try {
                    Write-Host "[-] connecting to graph... " -NoNewline
                    $TmpDelResponse = Invoke-RestMethod "https://graph.microsoft.com/beta/sites/$SiteId/permissions/$ApplicationId" -Method 'DELETE' -Headers $TmpDelHeaders
                    $TmpDelResponse | ConvertTo-Json
                    Write-Host "Success!" -ForegroundColor Green
                }
                catch {
                    Write-Host "Totes Fail!" -ForegroundColor Red
                    Write-Host "[-] $($Error[0])"
                }
            } else {
                Write-Host "[-] No application permissions"
            }
        }
        
        end {
            return $TmpDelResponse
        }
    }
    function Set-SPOSitePermissions {
        [CmdletBinding()]
        param (
            [Parameter()]
            [string]
            $SiteId,    
            [Parameter()]
            [string]
            $AppId,
            [Parameter()]
            [string]
            $DisplayName,
            [Parameter()]
            [string]
            $Role,
            [Parameter()]
            [string]
            $BearerToken
        )
        
        begin {
            Write-Host "[+] Adding App Registration to Site" -ForegroundColor Green
            if (($null -eq $AppId) -or ('' -eq $AppId)) {$AppId = Read-Host "[*] What is the App Id?"}
            if (($null -eq $DisplayName) -or ('' -eq $DisplayName)) {$DisplayName = Read-Host "[*] What is the App DisplayName?"}
            if (($null -eq $Role) -or ('' -eq $Role)) {$Role = Read-Host "[*] What role should the app have (read, write, owner)?"}
            $TmpHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $TmpHeaders.Add("Content-Type", "application/json")
            $TmpHeaders.Add("Authorization", "Bearer $BearerToken")
            $TmpBody = "
                {
                    `"roles`": [
                        `"$Role`"
                    ]`,
                    `"grantedToIdentities`": [
                        {
                            `"application`": {
                                `"id`": `"$AppId`"`,
                                `"displayName`": `"$DisplayName`"
                            }
                        }
                    ]
                }
            "
        }
        
        process {
            Write-Host "[-] connecting to graph... " -NoNewline
            try {
                $TmpResponse = Invoke-RestMethod "https://graph.microsoft.com/beta/sites/$SiteID/permissions" -Method 'POST' -Headers $TmpHeaders -Body $TmpBody
                Write-Host "Success!" -ForegroundColor Green
            }
            catch {
                Write-Host "Totes Fail!" -ForegroundColor Red
                Write-Host "[-] $($Error[0])"
            }
        }
        
        end {
            return $TmpResponse
        }
    }
    Show-Banner
    $BearerToken = Get-BearerToken -TenantId $TenantId -ConnectionAppId $ConnectionAppId -ConnectionSecret $ConnectionSecret
    $SPOSiteId = Get-SPOSiteID -SiteUrl $SiteUrl -BearerToken $BearerToken
    Get-SPOSitePermissions -SiteId $SPOSiteId -BearerToken $BearerToken
    if ($Add) {
        $TmpResponse = Set-SPOSitePermissions -SiteId $SPOSiteId -AppId $AppId -DisplayName $DisplayName -Role $Role -BearerToken $BearerToken
        Write-Host "[-] Updated Permission"
        Get-SPOSitePermissions -SiteId $SPOSiteId -BearerToken $BearerToken
    }
    if ($Delete) {
        $TmpDelResponse = Remove-SPOSitePermissions -SiteId $SPOSiteId -BearerToken $BearerToken
        Write-Host "[-] Updated Permission"
        Get-SPOSitePermissions -SiteId $SPOSiteId -BearerToken $BearerToken
    }
}

end {
    Write-Host "[-] Exiting"
}