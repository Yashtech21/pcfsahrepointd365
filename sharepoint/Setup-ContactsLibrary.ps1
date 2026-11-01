<#
.SYNOPSIS
    Provisions the "Contacts" document library columns/indexes and grants the backend
    app registration Sites.Selected write access to the AdvisorManagement site.

.REQUIREMENTS
    PnP.PowerShell 2.x  (Install-Module PnP.PowerShell -Scope CurrentUser)
    Run as a SharePoint admin. -PnPClientId is the Entra app you use for PnP interactive login.

.EXAMPLE
    ./Setup-ContactsLibrary.ps1 -SiteUrl https://contoso.sharepoint.com/sites/AdvisorManagement `
        -PnPClientId <pnp-login-app-id> -BackendAppId <D365-SP-DocManager app id>
#>
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [string] $PnPClientId,
    [Parameter(Mandatory)] [string] $BackendAppId,
    [string] $BackendAppName = "D365-SP-DocManager-Service",
    [string] $LibraryName = "Contacts"
)
$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $PnPClientId

# 1. Library
if (-not (Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue)) {
    New-PnPList -Title $LibraryName -Template DocumentLibrary -OnQuickLaunch | Out-Null
    Write-Host "Created library $LibraryName"
}

# 2. Columns (spec section 4.2)
$choices = @("Offer Letter", "Commissions", "Payslip", "Resignation Letter", "FinCEN_U4", "Misc")
$columns = @(
    @{ Name = "DocCategory";    Display = "Document Category"; Type = "Choice"; Indexed = $true  },
    @{ Name = "SourceEntity";   Display = "Source Entity";     Type = "Text";   Indexed = $false },
    @{ Name = "SourceRecordId"; Display = "Source Record ID";  Type = "Text";   Indexed = $true  },
    @{ Name = "ContactGuid";    Display = "Contact ID";        Type = "Text";   Indexed = $true  }
)
foreach ($c in $columns) {
    $field = Get-PnPField -List $LibraryName -Identity $c.Name -ErrorAction SilentlyContinue
    if (-not $field) {
        if ($c.Type -eq "Choice") {
            Add-PnPField -List $LibraryName -InternalName $c.Name -DisplayName $c.Display -Type Choice -Choices $choices -AddToDefaultView | Out-Null
        } else {
            Add-PnPField -List $LibraryName -InternalName $c.Name -DisplayName $c.Display -Type Text -AddToDefaultView | Out-Null
        }
        Write-Host "Added column $($c.Name)"
    }
    # Required columns would leave uploaded files checked out (the file is created before its metadata is set).
    $values = @{ Required = $false }
    if ($c.Indexed) { $values.Indexed = $true }
    Set-PnPField -List $LibraryName -Identity $c.Name -Values $values | Out-Null
    Write-Host "Configured column $($c.Name) (Required=false$(if ($c.Indexed) { ', Indexed=true' }))"
}

# Library settings that break the control. Warn rather than silently change a policy someone may have set.
$list = Get-PnPList -Identity $LibraryName
Get-PnPProperty -ClientObject $list -Property ForceCheckout, EnableModeration | Out-Null
if ($list.ForceCheckout) { Write-Warning "'Require check out' is ON for $LibraryName. Turn it off: uploaded files would stay checked out." }
if ($list.EnableModeration) { Write-Warning "'Require content approval' is ON for $LibraryName. Turn it off: uploads would stay Pending." }

# 3. Sites.Selected grant for the backend service principal (spec section 4.3)
Grant-PnPAzureADAppSitePermission -AppId $BackendAppId -DisplayName $BackendAppName -Site $SiteUrl -Permissions Write | Out-Null
Write-Host "Granted Write on $SiteUrl to $BackendAppName ($BackendAppId)"
