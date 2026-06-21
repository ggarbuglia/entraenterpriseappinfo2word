Clear-Host;

<#
.SYNOPSIS
Sets a built-in document property on a Word document.

.DESCRIPTION
Uses the Word COM object to set a built-in document property (e.g. Title, Subject, Keywords)
for the provided document. This is useful for adding metadata to generated Word files.

.PARAMETER Document
The Word Document COM object where the property will be set.

.PARAMETER Attribute
The name of the built-in property to set (for example: "Title", "Subject").

.PARAMETER Text
The text value to assign to the property.

.EXAMPLE
Set-BuiltInDocumentProperty -Document $doc -Attribute "Title" -Text "App Report"
#>
function Set-BuiltInDocumentProperty {
    param (
        [Parameter(Mandatory)] $Document,
        [Parameter(Mandatory)] [string] $Attribute,
        [Parameter(Mandatory)] [string] $Text
    )
    
    $obj = $Document.BuiltInDocumentProperties($Attribute)
    [System.__ComObject].InvokeMember("Value", [System.Reflection.BindingFlags]::SetProperty, $null, $obj, $Text)
}

<#
.SYNOPSIS
Adds a paragraph to a Word document with optional style.

.DESCRIPTION
Inserts a new paragraph at the end of the document content with the provided text
and applies the specified paragraph style. Useful for headings and body text.

.PARAMETER Document
The Word Document COM object where the paragraph will be added.

.PARAMETER Text
The paragraph text. Empty strings are allowed.

.PARAMETER Style
The name of the Word paragraph style to apply (defaults to 'Normal').

.EXAMPLE
Add-WordParagraph -Document $doc -Text 'Title' -Style 'Heading 1'
#>
function Add-WordParagraph {
    param(
        [Parameter(Mandatory)] $Document,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Text,
        [string] $Style = 'Normal'
    )

    $paragraph = $Document.Content.Paragraphs.Add()
    $paragraph.Range.Text = $Text
    $paragraph.Range.Style = $Style
    $paragraph.Range.InsertParagraphAfter() | Out-Null
}

<#
.SYNOPSIS
Adds a table with headers and rows to a Word document.

.DESCRIPTION
Creates a table in the document using the Word COM API. Accepts an array of header
strings and an array of row arrays (each row is an array of cell values). If there
are no rows, the function returns without modifying the document. The table style
is set to 'Table Grid' and the header row is bolded.

.PARAMETER Document
The Word Document COM object where the table will be added.

.PARAMETER Headers
An array of strings representing the column headers.

.PARAMETER Rows
An array of arrays; each inner array represents a row of cell values.

.EXAMPLE
$rows = @(); $rows += ,@('API','claim','True','False')
Add-WordTable -Document $doc -Headers @('API','Claim','Delegado','Consentimiento') -Rows $rows
#>
function Add-WordTable {
    param(
        [Parameter(Mandatory)] $Document,
        [Parameter(Mandatory)] [array] $Headers,
        [Parameter(Mandatory)] [array] $Rows
    )

    if ($Rows.Count -eq 0) {
        return
    }

    $table = $Document.Tables.Add($Document.Content.Paragraphs.Add().Range, $Rows.Count + 1, $Headers.Count)
    for ($col = 1; $col -le $Headers.Count; $col++) {
        $table.Cell(1, $col).Range.Text = $Headers[$col - 1]
    }

    for ($row = 0; $row -lt $Rows.Count; $row++) {
        for ($col = 0; $col -lt $Headers.Count; $col++) {
            if ($null -ne ($Rows[$row][$col])) {
                $table.Cell($row + 2, $col + 1).Range.Text = ($Rows[$row][$col]).ToString();
            }
            else {
                $table.Cell($row + 2, $col + 1).Range.Text = '';
            }
        }
    }

    $table.Style = 'Table Grid'
    $table.Rows.Item(1).Range.Bold = $true
    $table.AutoFitBehavior(2)
}

# =======================================================================================
# Entry Point
# =======================================================================================

# Connect to Microsoft Graph with the required scopes
Connect-MgGraph -NoWelcome -Scopes "Application.Read.All", "Directory.Read.All"

# Set output folder
$outputFolder = "C:\Temp\EntraAppsDocs"
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

# Get all Enterprise Applications (Service Principals)
$apps = Get-MgServicePrincipal -All #| Where-Object { $_.DisplayName -like "Azure*" }

# For each application (sorted by name)
foreach ($app in $apps | Sort-Object DisplayName) {

    Write-Host $app.DisplayName -ForegroundColor White

    # Get the corresponding application registration to access secrets, certs, and permissions
    $appReg = Get-MgApplication -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    if ($null -eq $appReg) { continue }

    # =======================================================================================
    # Extract secrets
    # =======================================================================================
    $secrets = @()
    if ($appReg.PasswordCredentials) {
        foreach ($secret in $appReg.PasswordCredentials) {
            $status = if ($secret.EndDateTime -gt (Get-Date)) {
                ""
            }
            else {
                "[Expired]"
            }
            $secrets += "Secret: $($secret.DisplayName) | Expiration: $($secret.EndDateTime.ToString("yyyy-MM-dd")) $status"
        }
    }

    # =======================================================================================
    # Extract certificates
    # =======================================================================================
    $certs = @()
    if ($appReg.KeyCredentials) {
        foreach ($cert in $appReg.KeyCredentials) {
            $certs += "Thumbprint: $($cert.KeyId) | Expiry: $($cert.EndDateTime.ToString("yyyy-MM-dd"))"
        }
    }

    # =======================================================================================
    # Extract permissions (API permissions)
    # =======================================================================================
    $permissions = @()
    if ($appReg.RequiredResourceAccess) {
        foreach ($perm in $appReg.RequiredResourceAccess) {
            # Try to resolve the resource (API) display name and permission definitions
            $resourceSp = Get-MgServicePrincipal -Filter "appId eq '$($perm.ResourceAppId)'" -ErrorAction SilentlyContinue
            if (-not $resourceSp) {
                # fallback: try to find by resourceAppId among all service principals
                $resourceSp = (Get-MgServicePrincipal -All | Where-Object { $_.AppId -eq $perm.ResourceAppId }) | Select-Object -First 1
            }

            foreach ($p in $perm.ResourceAccess) {
                $entry = $null
                $isDelegated = $false
                $desc = ''
                $name = ''
                $value = ''

                if ($p.Type -eq 'Role') {
                    $entry = $resourceSp.AppRoles | Where-Object { $_.Id -eq $p.Id }
                    if ($entry) {
                        $name = $entry.DisplayName
                        $value = $entry.Value
                        $desc = $entry.Description
                        $isDelegated = $false
                    }
                }
                elseif ($p.Type -eq 'Scope' -or $p.Type -eq 'Scope') {
                    $entry = $resourceSp.Oauth2PermissionScopes | Where-Object { $_.Id -eq $p.Id }
                    if ($entry) {
                        # oauth2 scopes expose admin consent display/description and value
                        $name = if ($entry.AdminConsentDisplayName) { $entry.AdminConsentDisplayName } else { $entry.Value }
                        $value = $entry.Value
                        $desc = if ($entry.AdminConsentDescription) { $entry.AdminConsentDescription } else { $entry.Description }
                        $isDelegated = $true
                    }
                }

                # Determine admin consent requirement conservatively
                $adminConsentRequired = $false
                if ($p.Type -eq 'Role') { $adminConsentRequired = $true }
                elseif ($entry -and ($entry.AdminConsentDescription -or $entry.AdminConsentDisplayName)) { $adminConsentRequired = $true }

                $permissions += [PSCustomObject]@{
                    API                  = if ($resourceSp) { $resourceSp.DisplayName } else { $perm.ResourceAppId }
                    Claim                = $value
                    Permission           = $name
                    Description          = $desc
                    IsDelegated          = $isDelegated
                    AdminConsentRequired = $adminConsentRequired
                }
            }
        }
    }

    # =======================================================================================
    # Extract custom roles defined in the application registration
    # =======================================================================================
    $customRoles = @()
    if ($appReg.AppRoles) {
        foreach ($role in $appReg.AppRoles | Where-Object { $_.DisplayName -ne 'msiam_access' }) {
            $customRoles += [PSCustomObject]@{
                Name        = $role.DisplayName
                Value       = $role.Value
                Id          = $role.Id
                Description = $role.Description
            }
        }
    }

    # =======================================================================================
    # Extract app owners
    # =======================================================================================
    $owners = @()
    try {
        $appOwners = Get-MgServicePrincipalOwner -ServicePrincipalId $app.Id -ErrorAction SilentlyContinue
        if ($appOwners) {
            foreach ($owner in $appOwners) {
                $owners += $owner.AdditionalProperties.displayName
            }
        }
    }
    catch {
        Write-Host "Warning: Could not retrieve owners for $($app.DisplayName)" -ForegroundColor Yellow
    }

    # =======================================================================================
    # Extract SAML-specific configuration
    # =======================================================================================
    $samlConfig = [PSCustomObject]@{
        EntityIds             = @()
        AcsUrls               = @()
        LoginUrl              = ''
        NotificationEmails    = @()
        CertificateThumbprint = ''
        CertificateDetails    = @()
        ClaimMappingPolicies  = @()
    }

    if ($app.PreferredSingleSignOnMode -eq 'saml') {
        $samlSp = Get-MgServicePrincipal -ServicePrincipalId $app.Id -Property PreferredSingleSignOnMode, LoginUrl, ReplyUrls, ServicePrincipalNames, NotificationEmailAddresses, PreferredTokenSigningKeyThumbprint, KeyCredentials -ErrorAction SilentlyContinue
        if ($samlSp) {
            if ($samlSp.ServicePrincipalNames) { $samlConfig.EntityIds = $samlSp.ServicePrincipalNames }
            if ($samlSp.ReplyUrls) { $samlConfig.AcsUrls = $samlSp.ReplyUrls }
            $samlConfig.LoginUrl = $samlSp.LoginUrl
            if ($samlSp.NotificationEmailAddresses) { $samlConfig.NotificationEmails = $samlSp.NotificationEmailAddresses }
            $samlConfig.CertificateThumbprint = $samlSp.PreferredTokenSigningKeyThumbprint
            if ($samlSp.KeyCredentials) {
                foreach ($cred in $samlSp.KeyCredentials) {
                    $samlConfig.CertificateDetails += [PSCustomObject]@{
                        KeyId         = $cred.KeyId
                        DisplayName   = $cred.DisplayName
                        StartDateTime = $cred.StartDateTime
                        EndDateTime   = $cred.EndDateTime
                    }
                }
            }

            try {
                $claimPolicies = Get-MgServicePrincipalClaimMappingPolicy -ServicePrincipalId $app.Id -ErrorAction SilentlyContinue
                if ($claimPolicies) {
                    foreach ($policy in $claimPolicies) {
                        $claims = $policy.AdditionalProperties['claimsMapping'] ?? $policy.AdditionalProperties['claimMapping'] ?? $null
                        $samlConfig.ClaimMappingPolicies += [PSCustomObject]@{
                            Id          = $policy.Id
                            DisplayName = $policy.DisplayName
                            Description = $policy.Description
                            Claims      = $claims
                        }
                    }
                }
            }
            catch {
                # ignore claim mapping failures
            }
        }
    }

    # =======================================================================================
    # Extract assignments to users/groups/roles (app role assignments)
    # =======================================================================================
    $assignedList = @();
    $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $app.Id -All
    $assignments | ForEach-Object {
        $roleDisplay = ''

        if ($_.AppRoleId -and $appReg.AppRoles) {
            $roleEntry = $appReg.AppRoles | Where-Object { $_.Id -eq $_.AppRoleId }
            if ($roleEntry) { $roleDisplay = $roleEntry.DisplayName } else { $roleDisplay = $_.AppRoleId }
        }

        $assignedList += [PSCustomObject]@{
            Name = $_.PrincipalDisplayName
            Type = $_.PrincipalType
            Role = $roleDisplay
        }
    }

    $provisioning = if ($app.ServicePrincipalType -eq "EnterpriseApplication") { "Possible / Check Portal" } else { "N/A" }

    $authType = if ($app.PreferredSingleSignOnMode) {
        $app.PreferredSingleSignOnMode
    }
    else {
        "OAuth/OIDC"
    }

    # =======================================================================================
    # Process Word document generation using COM objects
    # =======================================================================================

    $tmpPath = "$outputFolder\template.dotx"
    $docPath = "$outputFolder\$($app.DisplayName -replace '[^a-zA-Z0-9]', '_').docx"
    $docPath = $docPath -replace '__', '_';
    $docPath = $docPath -replace '__', '_';

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add($tmpPath)

    Set-BuiltInDocumentProperty -Document $doc -Attribute "Title" -Text "Entra Enterprise App: $($app.DisplayName)"
    Set-BuiltInDocumentProperty -Document $doc -Attribute "Subject" -Text "Enterprise App '$($app.DisplayName)' configuration definition."
    Set-BuiltInDocumentProperty -Document $doc -Attribute "Keywords" -Text "Entra;ID;Enterprise;App;Registration;Azure"

    try {
        # =======================================================================================
        # Application Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text 'Indetifiers' -Style 'Heading 1'
        $multilineText = "";
        $multilineText += "Tenant ID: $((Get-MgContext).TenantId)" + [char]11
        $multilineText += "Application ID: $($app.AppId)" + [char]11
        $multilineText += "Objet ID: $($app.Id)"
        Add-WordParagraph -Document $doc -Text $multilineText

        # =======================================================================================
        # Application Owners
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'Owners' -Style 'Heading 1'
        if ($owners.Count -gt 0) {
            $multilineText = ($owners | Sort-Object) -join [char]11
            Add-WordParagraph -Document $doc -Text $multilineText
        }
        else {
            Add-WordParagraph -Document $doc -Text 'No owners found.'
        }

        # =======================================================================================
        # Application Authentication Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'Authentication' -Style 'Heading 1'
        if ($app.PreferredSingleSignOnMode -eq 'saml') {
            $multilineText = "";
            $multilineText += "SSO Type: $($authType.ToUpper())" + [char]11
            $multilineText += "Identifier (Entity ID):" + [char]11
            if ($samlConfig.EntityIds.Count -gt 0) {
                $multilineText += ($samlConfig.EntityIds | Where-Object { $_ -like "http*" } | Sort-Object) -join [char]11
            }
            else {
                $multilineText += 'No Entity ID identifier found.' + [char]11
            }

            $multilineText += '' + [char]11
            $multilineText += 'Sign on URL (Assertion Consumer Service URLs):' + [char]11
            if ($samlConfig.AcsUrls.Count -gt 0) {
                $multilineText += ($samlConfig.AcsUrls | Sort-Object) -join [char]11
            }
            else {
                $multilineText += 'No Assertion Consumer Service URL found.' + [char]11
            }

            $multilineText += '' + [char]11
            $multilineText += "Sign on URL:" + [char]11
            $multilineText += "$($samlConfig.LoginUrl)" + [char]11

            $multilineText += '' + [char]11
            $multilineText += 'Notification Email:' + [char]11
            if ($samlConfig.NotificationEmails.Count -gt 0) {
                $multilineText += ($samlConfig.NotificationEmails | Sort-Object) -join [char]11
            }
            else {
                $multilineText += 'No notification email found.' + [char]11
            }

            $multilineText += '' + [char]11
            $multilineText += "SAML Certificate Thumbprint:" + [char]11
            $multilineText += "$($samlConfig.CertificateThumbprint)" + [char]11

            if ($samlConfig.CertificateDetails.Count -gt 0) {
                $multilineText += '' + [char]11
                $multilineText += 'SAML Certificate Details:' + [char]11
                foreach ($cert in $samlConfig.CertificateDetails) {
                    $multilineText += '' + [char]11
                    $multilineText += "Name: $($cert.DisplayName)" + [char]11
                    $multilineText += "KeyId: $($cert.KeyId)" + [char]11
                    $multilineText += "Expiration: $($cert.EndDateTime)" + [char]11
                }
            }

            Add-WordParagraph -Document $doc -Text $multilineText
        }
        elseif ($appReg.Web.RedirectUris.Count -gt 0) {
            $multilineText = "";
            $multilineText += "SSO Type: $($authType.ToUpper())" + [char]11
            $multilineText += "Redirection URIs:" + [char]11
            $multilineText += ($appReg.Web.RedirectUris | Sort-Object) -join [char]11
            Add-WordParagraph -Document $doc -Text $multilineText
        }
        else {
            Add-WordParagraph -Document $doc -Text "No Redirection URIs found."
        }

        # =======================================================================================
        # Application Certificate and Secret Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'Certificates & Secrets' -Style 'Heading 1'
        if ($certs.Count -gt 0) {
            $multilineText = ($certs | Sort-Object) -join [char]11
            Add-WordParagraph -Document $doc -Text $multilineText   
        }
        else {
            Add-WordParagraph -Document $doc -Text 'No certificates found.'
        }

        if ($secrets.Count -gt 0) {
            $multilineText = ($secrets | Sort-Object) -join [char]11
            Add-WordParagraph -Document $doc -Text $multilineText   
        }
        else {
            Add-WordParagraph -Document $doc -Text 'No secrets found.'
        }

        # =======================================================================================
        # Application Assignments Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'Assignments' -Style 'Heading 1'
        if ($assignedList.Count -gt 0) {
            $rows = @()
            foreach ($assignment in $assignedList) {
                $rows += , @($assignment.Name, $assignment.Type, $assignment.Role)
            }
            Add-WordTable -Document $doc -Headers @('DisplayName', 'Type', 'Role') -Rows $rows
        }
        else {
            Add-WordParagraph -Document $doc -Text 'No assignments found.'
        }

        # =======================================================================================
        # Application Provisioning Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'Provisioning' -Style 'Heading 1'
        Add-WordParagraph -Document $doc -Text "Provisioning Method: $provisioning"

        # =======================================================================================
        # Application API Permissions Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'API Permissions' -Style 'Heading 1'
        if ($permissions.Count -gt 0) {
            $rows = @()
            foreach ($perm in $permissions) {
                $rows += , @($perm.API, $perm.Claim, $perm.IsDelegated, $perm.AdminConsentRequired)
            }
            Add-WordTable -Document $doc -Headers @('API', 'Claim', 'Delegated', 'Consent') -Rows $rows
        }
        else {
            Add-WordParagraph -Document $doc -Text 'No API permissions found.'
        }

        # =======================================================================================
        # Application Custom Roles Information
        # =======================================================================================
        Add-WordParagraph -Document $doc -Text '' -Style 'Normal'
        Add-WordParagraph -Document $doc -Text 'Custom Roles' -Style 'Heading 1'
        if ($customRoles.Count -gt 0) {
            $rows = @()
            foreach ($role in $customRoles) {
                $rows += , @($role.Name, $role.Value, $role.Description)
            }
            Add-WordTable -Document $doc -Headers @('Name', 'Value', 'Description') -Rows $rows
        }
        else {
            Add-WordParagraph -Document $doc -Text 'No custom roles found.'
        }

        $doc.SaveAs2($docPath, 16)
    }
    finally {
        $doc.Close($false)
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    Write-Host "--- Generated: $docPath" -ForegroundColor Green
}

Write-Host;
Write-Host 'Completed documentation generation with COM' -ForegroundColor Cyan
