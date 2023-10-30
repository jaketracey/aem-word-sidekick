# AEM Document Based Authoring Word Add in

This is an addin that adds a Word sidebar that has similar functionality to the AEM Sidekick.

It allows you to setup your AEM edge delivery site and then preview and publish directly from the Word interface.

## Usage

You can sideload the manifest.xml from the dist folder into Microsoft Word for the latest version, or install it from AppSource.


### Deploy to Word desktop app
```
npm run start
```

### Deploy to Word Online
```
npm run start:web -- --document 'link-to-sharepoint-doc'
```