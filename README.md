# Cisco Network Registrar Importer

The goal of this script was to be able to take XLSX data and import that data into Network Registrar in mass.

## Getting Started

Look at the template XLSX. The column names are important in the FinalConfig page. Use the OriginalConfig tab to fill data into the FinalData tab.

Step 1. Build the template (Original Data and Config are used to programmatically fill data on the FinalData tab. All data in the script is pulled from FinalData)
Step 3. Run and profit

Report any issues to my email and I will get them fixed.

### Prerequisites

GIT (This is required to download the XLHELPER module using a fork that  I made for compatibility with Python 2.7)
XLHELPER
OPENPYXL

## Deployment

Just execute the script and answer the questions

## Features
- XLSX-based import
- Export directly into Network Registar via API
- Export XML files to disk

## *Caveats
- None

## Versioning

VERSION 1.0
Currently Implemented Features
- XLSX-based import
- Export directly into Network Registar via API
- Export XML files to disk

## Authors

* **Matt Cross** - [RouteAllThings](https://github.com/routeallthings)
* **Ben Wantrobski**

See also the list of [contributors](https://github.com/routeallthings/NR-Importer/contributors) who participated in this project.

## License

This project is licensed under the GNU - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Thanks to HBS for giving me a reason to write this.
