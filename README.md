# Entra Enterprise App Information to Word Document

Azure Entra ID Enterprise Apps information extractor to Word documents using a Word template and Microsoft Graph.

## Overview

`Get-EntraEnterpriseAppsInfo.ps1` connects to Microsoft Graph, retrieves enterprise application data from Microsoft Entra ID, and generates a Word document for each application.

The script uses Word COM automation to populate a `.docx` report from `template.dotx`.

## Included files

- `Get-EntraEnterpriseAppsInfo.ps1` - main PowerShell script that collects app data and generates reports.
- `template.dotx` - Word document template used to create each report.
- `README.md` - this documentation.

## Prerequisites

- Windows with Microsoft Word installed.
- PowerShell and the Microsoft Graph PowerShell SDK.
- Access to Microsoft Entra ID with permissions to read applications and service principals.
- `template.dotx` must exist in the output folder, as the script expects it at `$outputFolder\template.dotx`.

## Required Graph permissions

The script requests the following Microsoft Graph scopes:

- `Application.Read.All`
- `Directory.Read.All`

## How to run

1. Open PowerShell in the repository directory.
2. Run the script:

```powershell
.\nGet-EntraEnterpriseAppsInfo.ps1
```

The script will prompt for Microsoft Graph sign-in if needed.

## Output

- Generated reports are saved to `C:\Temp\EntraAppsDocs`.
- Output files use the application display name, with invalid filename characters replaced by underscores.
- Example output file: `C:\Temp\EntraAppsDocs\My_Enterprise_App.docx`.

## What the script collects

For each enterprise application, the script extracts:

- Tenant ID, Application ID, and Object ID
- Application owners
- Authentication details:
  - SAML entity IDs, ACS URLs, login URL, notification emails, certificate details
  - OAuth/OIDC redirect URIs when available
- Certificates and secrets
- App role assignment recipients
- Provisioning method indicator
- API permissions and consent requirements
- Custom application roles

## Document generation details

The script uses these helper functions:

- `Set-BuiltInDocumentProperty` - writes built-in Word metadata like Title, Subject, and Keywords.
- `Add-WordParagraph` - appends paragraphs with optional styles.
- `Add-WordTable` - inserts tables with header rows and cell values.

It starts Word via COM, opens the template, inserts report sections, saves the file, and releases COM objects.

## Notes

- Word COM automation requires a local installation of Microsoft Word.
- The output folder and template path are hard-coded in the script.
- If a section has no data, the generated document explicitly notes that no values were found.
- The script currently processes all service principals returned by `Get-MgServicePrincipal -All`.

## Troubleshooting

- If Word cannot start, verify that Microsoft Word is installed and COM automation is supported.
- If Graph permissions are missing, grant the required scopes or run as an account with sufficient access.
- If the template is not found, ensure `template.dotx` exists in `C:\Temp\EntraAppsDocs` before running the script.

