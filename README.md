# spfx-list-charts

An SharePoint Framework web part that uses the Chart.js library to visual SharePoint list data.

## Built in Chart Types

- Bar
- Horzontal Bar
- Doughnut
- Line
- Pie

## Themes

Each chart can be uniquely themed with the built-in color theme generator (color-scheme), continue generating a theme until you find one that matches your style.

New charts are populated with Sample data by default. Select a list data source, label column, data column and which column indicates a unique value in your list and your chart will generate dynamically.

## Current Data Functions

- Average
- Count
- Sum

### Working With

- [SharePoint SPFx](https://docs.microsoft.com/en-us/sharepoint/dev)
- [Office Graph](https://developer.microsoft.com/en-us/graph/docs/concepts/get-started)
- [React](https://reactjs.org)
- [Chart.JS](https://www.chartjs.org/)

### Applies to

- [SharePoint Framework Developer Preview](http://dev.office.com/sharepoint/docs/spfx/sharepoint-framework-overview)
- [Office 365 developer tenant](http://dev.office.com/sharepoint/docs/spfx/set-up-your-developer-tenant)

### Solution

Solution|Author(s)
--------|---------
spfx-list-chart|Anthony Conrad

Version|Date|Comments
-------|----|--------
0.0.1.0|April 5, 2018|Initial Release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

### Building the code

```bash
git clone https://github.com/parithon/spfx-list-chart.git
npm i
npm i -g gulp
gulp
```

This package produces the following:

- lib/* - intermediate-stage commonjs build artifacts
- dist/* - the bundled script, along with other resources
- deploy/* - all resources which should be uploaded to a CDN.
