<#
.SYNOPSIS
    Read-only check that the SharePoint side matches docs/PREREQUISITES.md §3.
    Prints PASS/FAIL per item and exits with code 1 if anything failed.

.REQUIREMENTS
    PowerShell 7.2+, PnP.PowerShell 2.x. Needs read access to the site; the grant check needs a SharePoint admin.

.EXAMPLE
    ./Test-ContactsLibrary.ps1 -SiteUrl https://contoso.sharepoint.com/sites/AdvisorManagement `
        -PnPClientId <pnp-login-app-id> -BackendAppId <D365-SP-DocManager app id>
#>
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [Parameter(Mandatory)] [string] $PnPClientId,
    [string] $BackendAppId,
    [string] $LibraryName = "Contacts"
)
$ErrorActionPreference = "Stop"
$script:failures = 0

function Report([bool] $ok, [string] $message) {
    if ($ok) { Write-Host "PASS  $message" -ForegroundColor Green }
    else { Write-Host "FAIL  $message" -ForegroundColor Red; $script:failures++ }
}

Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $PnPClientId

# Library
$list = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue
Report ($null -ne $list) "Library '$LibraryName' exists"
if (-not $list) { exit 1 }
Report ($list.BaseTemplate -eq 101) "'$LibraryName' is a document library"

Get-PnPProperty -ClientObject $list -Property ForceCheckout, EnableModeration | Out-Null
Report (-not $list.ForceCheckout) "'Require check out' is off"
Report (-not $list.EnableModeration) "'Require content approval' is off"

# Columns (internal names, types, index, not required)
$expected = @(
    @{ Name = "DocCategory";    Type = "Text";   Indexed = $true  },
    @{ Name = "SourceEntity";   Type = "Text";   Indexed = $false },
    @{ Name = "SourceRecordId"; Type = "Text";   Indexed = $true  },
    @{ Name = "ContactGuid";    Type = "Text";   Indexed = $true  }
)
foreach ($c in $expected) {
    $f = Get-PnPField -List $LibraryName -Identity $c.Name -ErrorAction SilentlyContinue
    if (-not $f -or $f.InternalName -ne $c.Name) {
        Report $false "Column with internal name '$($c.Name)' exists (create it with this internal name, then rename the display name)"
        continue
    }
    Report ($f.TypeAsString -eq $c.Type) "$($c.Name) is type $($c.Type) (found $($f.TypeAsString)$(if ($f.TypeAsString -eq 'Choice') { '; change it to Single line of text' }))"
    Report (-not $f.Required) "$($c.Name) is not Required"
    if ($c.Indexed) { Report ([bool]$f.Indexed) "$($c.Name) is indexed" }
}

# DocCategory is text: category values come from Dataverse (ce_documentcategory), so there is no
# choice list to compare. A Choice column would reject categories added later in Dataverse.

# Sites.Selected grant for the backend app
if ($BackendAppId) {
    try {
        $grants = Get-PnPAzureADAppSitePermission -Site $SiteUrl -AppIdentity $BackendAppId -ErrorAction Stop
        $roles = @($grants | ForEach-Object { $_.Roles }) | ForEach-Object { "$_".ToLower() }
        Report (($roles -contains "write") -or ($roles -contains "fullcontrol") -or ($roles -contains "owner")) `
            "Backend app $BackendAppId has Write on the site (found: $($roles -join ', '))"
    } catch {
        Report $false "Could not read the Sites.Selected grant for $BackendAppId (needs a SharePoint admin): $($_.Exception.Message)"
    }
}

if ($script:failures -gt 0) {
    Write-Host "`n$($script:failures) check(s) failed. See docs/PREREQUISITES.md section 3." -ForegroundColor Red
    exit 1
}
Write-Host "`nAll SharePoint prerequisites are in place." -ForegroundColor Green
