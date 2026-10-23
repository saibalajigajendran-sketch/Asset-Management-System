# Asset Management System

## Summary

The Asset Management System is a SharePoint Framework (SPFx) solution for managing organizational assets. It provides asset registration, attachment management, dashboard reporting, and laptop asset tracking.

## SharePoint Framework Version

* SPFx: 1.23.2
* React: 17
* TypeScript

## Features

### Asset Attachments

Asset files are stored as SharePoint list item attachments. The web part uses the authenticated SPFx `SPHttpClient` context to manage attachments without creating a separate document library.

Supported file types include PDF, JPG/JPEG, PNG, DOCX, XLSX, MSG, and EML. Each file can be up to 10 MB, with up to 10 files per asset.

To enable attachments on the existing asset registers, connect PnP PowerShell to the target SharePoint site and run:

```powershell
./scripts/Enable-AssetRegisterAttachments.ps1
```

Confirm that the configured SharePoint lists exist before running the script.

### Dashboard

The Dashboard is the default landing page and retrieves live information from configured SharePoint lists through the web part's authenticated `SPHttpClient`.

Register names, list URLs, field aliases, and stock thresholds are configured in:

`src/webparts/assetManagement/components/dashboardConfig.ts`

Dashboard results are cached in memory for five minutes. Use Refresh to bypass the cache.

### Laptop Asset Timeline

The Hardware - Laptops, Desktops View modal supports tab navigation and asset timeline information.

Timeline dates come from the selected SharePoint list item. Movement and service records are associated by SharePoint item ID to avoid mixing histories when Asset IDs are duplicated.

Timeline configuration is maintained in:

`src/webparts/assetManagement/components/assetTimelineConfig.ts`

Actual service-request history requires a configured SharePoint service-request list linked to the asset list item ID.

## Prerequisites

* A SharePoint Online environment.
* Appropriate permissions to access the configured SharePoint lists.
* Node.js compatible with the project's `package.json` requirements.
* npm.

## Build and Package

Install dependencies:

```bash
npm ci
```

Build and package the solution:

```bash
npm run build
```

The SharePoint package is expected at:

`sharepoint/solution/asset-management.sppkg`

## GitHub Actions Deployment

The repository includes a GitHub Actions workflow at:

`.github/workflows/deploy-spfx-solution.yml`

The workflow is configured to run when changes are pushed to the `main` branch and can also be triggered manually.

Deployment requires the GitHub repository secrets configured in the workflow and an appropriately configured Microsoft Entra application with the required permissions.

## Version History

| Version | Comments                                                                                    |
| ------- | ------------------------------------------------------------------------------------------- |
| 1.0     | Initial Asset Management System release                                                     |
| 1.1     | Added attachment management, dashboard enhancements, and laptop asset timeline improvements |

## Disclaimer

This code is provided as-is, without warranties of any kind. Review and test the solution in your environment before deploying it to production.

## References

* [SharePoint Framework documentation](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
* [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)
* [Heft documentation](https://heft.rushstack.io/)
