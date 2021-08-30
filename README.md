# spfx-extend-lang

## Summary

Extend language select to add link to account language settings.

Language toggles on Hub sites will be automatically be targeted by this extension. To target sub-sites of a hub site that also has language toggles, change the hubSiteId values on line 22 of [ExtendLanguageApplicationCustomizer](https://github.com/gcxchange-gcechange/spfx-extendlang/blob/main/src/extensions/extendLanguage/ExtendLanguageApplicationCustomizer.ts)

## Deployment

spfx-extendlang is intended to be deployed **tenant wide**

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**
