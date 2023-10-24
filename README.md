# AEM Document Based Authoring Word Add in

_This is a POC done over a couple of evenings to prove the concept, I'll completely refactor it shortly :)_

This addin adds a sidebar to the Microsoft Word user interface similar to the AEM Sidekick chrome extension.

It allows you to setup your AEM edge delivery site and then preview and publish directly from the Word interface.

TODO:
Pass checks to be able to deploy to AppSource
Refactor everything

## Usage

### Deploy to Word desktop app
```
npm run start
```

### Deploy to Word Online
```
npm run start:web -- --document 'link-to-sharepoint-doc'
```

