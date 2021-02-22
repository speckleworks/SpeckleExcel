# SpeckleExcel

[![Netlify Status](https://api.netlify.com/api/v1/badges/e16a8dbc-8084-42d2-aa09-a291b9284b59/deploy-status)](https://app.netlify.com/sites/speckleexcel/deploys)

Speckle client for Microsoft Excel

## ⚠️ IMPORTANT ⚠️

Speckle v1 has reached EOL and is no longer mainatined. Check out our brand new 👉 [Speckle v2](https://github.com/specklesystems)!
Read more about the announcemnt [here](https://speckle.systems/blog/speckle2-vision-and-faq) and check out or new [website](https://speckle.systems).


![SpeckleExcel](https://github.com/speckleworks/SpeckleExcel/raw/master/images/speckleexcel.png)

## Installation
SpeckleExcel will be added to the Office AppSource soon. For now, you can sideload the plugin using the following steps:
1. Download `manifest.xml` from the SpeckleExcel folder [here](https://raw.githubusercontent.com/speckleworks/SpeckleExcel/master/SpeckleExcel/manifest.xml)
2. Sideload the app:
	- [Sideload Office Add-ins in Windows](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
	- [Sideload Office Add-ins in Office on the web](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)


## Server requirements
- SpeckleExcel requires updated Speckle servers (minimum [commit](https://github.com/speckleworks/SpeckleServer/commit/9e135c453a93608a7e75d0317407070a64bdcea7) supported)
- Please ensure that your Speckle server has `https://excel.speckle.systems` specified under `REDIRECT_URL` within the `.env` file

## Build Setup

SpeckleExcel contains two Office add-ins to sideload, one for release and one for development:
- SpeckleExcel: uses `https://excel.speckle.systems` as the plugin source
- SpeckleExcelDev: uses `https://localhost:8080` as the plugin source
  - Make sure to add the certificate from `https://localhost:8080` before loading the plugin

### Build instructions

``` bash
# install dependencies
npm install

# serve with hot reload at localhost:8080
npm run start

# build for production with minification
npm run build
```
