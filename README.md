# asset-management

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install -g @rushstack/heft`
  - `npm install`
  - `heft start`

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

### Asset attachments

Asset files are stored as built-in SharePoint list item attachments. The web part does not create a document library or attachment metadata columns. Attachment operations use the SharePoint REST API through the existing `SPHttpClient` context. Each register's configured list must have attachments enabled, and the signed-in user needs permission to add, read, and delete list item attachments.

The upload accepts PDF, JPG/JPEG, PNG, DOCX, XLSX, MSG, and EML files up to 10 MB each, with up to 10 files per asset. Files keep their original names. Names that conflict with existing attachments or violate SharePoint filename and path rules are rejected inline. If a newly created asset has failed uploads, the item remains in SharePoint and the form offers retry for remaining files.

To enable attachments on the existing 12 registers, connect PnP PowerShell to the target site with `Connect-PnPOnline`, then run:

```powershell
./scripts/Enable-AssetRegisterAttachments.ps1
```

The script uses the 12 list titles configured in the web part. Confirm that these titles exist in the target SharePoint site before running it. If adding a register, add its list URL to `assetRegisterUrls` in `AssetManagement.tsx` and its key to `scripts/Enable-AssetRegisterAttachments.ps1`; the shared `<AssetAttachments />` component and attachment REST service are reused without register-specific file code. Attachment upload does not depend on an asset identifier.

### Dashboard

The Dashboard is the default landing page and reads live values from the configured SharePoint lists through the web part's authenticated `SPHttpClient`. Register names, list URLs, and field-name aliases are defined in `src/webparts/assetManagement/components/dashboardConfig.ts`; aliases are resolved against each list's SharePoint field metadata, and missing fields are skipped. When adding a register, add its URL and field aliases there. Configure per-register low-stock cutoffs in `dashboardStockThresholds` in the same file. Dashboard results are cached in memory for five minutes; use Refresh to bypass the cache.

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

### Laptop asset timeline

The Hardware - Laptops, Desktops View modal enables tab navigation with `timelineTab: true` in `dashboardConfig.ts`. Timeline dates come from the selected list item, and movement/service rows are associated by a SharePoint item ID reference so duplicate Asset IDs do not mix histories. Tracker field aliases are in `assetTimelineConfig.ts` and are matched against that list's live field metadata. To enable another register, add its timeline flag and source field aliases, then pass that register's purchase and support date aliases to the timeline view.

This repository did not contain a configured SharePoint service-request list; the previous service history was generated sample content and has been removed. To include actual service requests, configure `serviceRequestTimelineSource` in `assetTimelineConfig.ts` with the existing list URL and its item-ID, reported/resolved date, title, technician, and status field aliases. The source must link to the asset list item ID; Asset ID alone is intentionally not used. The app uses the existing SPFx `SPHttpClient` service and resolves each list's actual field types before selecting or expanding fields.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
